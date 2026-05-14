"""
SafeSales order SMS — WinUI 3 shell (win32more).

Run: ``python -m ui`` from the repository root.
"""

from __future__ import annotations

import threading
import traceback
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

from win32more.Microsoft.UI.Dispatching import DispatcherQueuePriority
from win32more.Microsoft.UI.Text import FontWeights
from win32more.Microsoft.UI.Xaml import (
    CornerRadiusHelper,
    ElementTheme,
    GridLengthHelper,
    GridUnitType,
    HorizontalAlignment,
    TextWrapping,
    Thickness,
    VerticalAlignment,
    Visibility,
    Window,
)
from win32more.Microsoft.UI.Xaml.Controls import (
    Border,
    Button,
    ColumnDefinition,
    ComboBox,
    ComboBoxItem,
    Grid,
    NavigationView,
    NavigationViewBackButtonVisible,
    NavigationViewItem,
    NavigationViewPaneDisplayMode,
    Orientation,
    ProgressRing,
    RowDefinition,
    ScrollBarVisibility,
    ScrollMode,
    ScrollViewer,
    StackPanel,
    Symbol,
    SymbolIcon,
    TextBlock,
    TextBox,
)
from win32more.Microsoft.UI.Xaml.Media import FontFamily, MicaBackdrop, SolidColorBrush
from win32more.Microsoft.UI.Xaml.Shapes import Ellipse
from win32more.Microsoft.Windows.Storage.Pickers import FileOpenPicker, PickFileResult, PickerLocationId
from win32more.Windows.Foundation import AsyncStatus, TimeSpan
from win32more.Windows.UI import Color
from win32more.winui3 import XamlApplication

from easysms import EasySMSError
from ui import services


def _ts() -> str:
    return datetime.now().strftime("%H:%M:%S")


def _rgb(a: int, r: int, g: int, b: int) -> SolidColorBrush:
    c = Color()
    c.A = a
    c.R = r
    c.G = g
    c.B = b
    return SolidColorBrush(c)


@dataclass(frozen=True, slots=True)
class _ChromePalette:
    surface: tuple[int, int, int]
    card: tuple[int, int, int]
    border: tuple[int, int, int]
    muted: tuple[int, int, int]
    primary: tuple[int, int, int]
    accent: tuple[int, int, int]
    online: tuple[int, int, int]
    offline: tuple[int, int, int]
    on_accent: tuple[int, int, int]


def _palette_light() -> _ChromePalette:
    return _ChromePalette(
        surface=(246, 246, 248),
        card=(252, 252, 253),
        border=(225, 226, 230),
        muted=(96, 98, 102),
        primary=(26, 28, 34),
        accent=(0, 103, 192),
        online=(16, 124, 16),
        offline=(196, 43, 28),
        on_accent=(255, 255, 255),
    )


def _palette_dark() -> _ChromePalette:
    return _ChromePalette(
        surface=(32, 32, 36),
        card=(43, 43, 47),
        border=(58, 58, 64),
        muted=(168, 168, 174),
        primary=(243, 243, 243),
        accent=(76, 156, 230),
        online=(108, 203, 95),
        offline=(255, 138, 128),
        on_accent=(255, 255, 255),
    )


class SafeSalesWinUIApp(XamlApplication):
    def OnLaunched(self, args):
        self._theme_mode: str = "system"
        self._theme_combo_guard = False
        self._themed_borders: list[Border] = []
        self._root: Grid | None = None
        self._log_wrap: Border | None = None
        self._log_header: TextBlock | None = None
        self._account_outer: Border | None = None
        self._theme_combo: ComboBox | None = None
        self._nav: NavigationView | None = None
        self._nav_item_sms: NavigationViewItem | None = None
        self._nav_item_orders: NavigationViewItem | None = None
        self._nav_item_settings: NavigationViewItem | None = None
        self._page_host: Grid | None = None
        self._page_sms: ScrollViewer | None = None
        self._page_orders: ScrollViewer | None = None
        self._page_settings: ScrollViewer | None = None
        self._current_route: str = "sms"
        self._nav_header_title: TextBlock | None = None
        self._nav_header_subtitle: TextBlock | None = None
        self._settings_hint: TextBlock | None = None
        self._send_btn: Button | None = None
        self._bulk_btn: Button | None = None

        self._window = Window()
        self._window.Title = "SafeSales · SMS"
        self._window.SystemBackdrop = MicaBackdrop()

        self._dispatcher = self._window.DispatcherQueue
        self._client = services.get_client()

        self._balance_text: TextBlock | None = None
        self._api_status_text: TextBlock | None = None
        self._api_dot: Ellipse | None = None
        self._api_online = False
        self._pulse_timer = None
        self._pulse_phase = True

        self._log_box: TextBox | None = None
        self._acs_path = ""
        self._box_path = ""
        self._acs_path_box: TextBox | None = None
        self._box_path_box: TextBox | None = None
        self._excel_results: TextBox | None = None
        self._excel_template: TextBox | None = None
        self._excel_sender: TextBox | None = None
        self._last_shipments: list = []
        self._progress: ProgressRing | None = None

        self._sms_phone: TextBox | None = None
        self._sms_message: TextBox | None = None
        self._sms_sender: TextBox | None = None
        self._sms_status: TextBlock | None = None

        self._pick_target: str | None = None

        self._themed_borders.clear()
        root = Grid()
        self._root = root
        root.RowDefinitions.Append(self._row_star())
        root.RowDefinitions.Append(self._row_px(180))

        page_sms = self._build_one_time_tab()
        page_orders = self._build_excel_tab()
        page_settings = self._build_settings_tab()
        self._page_sms = page_sms
        self._page_orders = page_orders
        self._page_settings = page_settings

        # Pages live inside a single host grid; we toggle Visibility to swap views.
        # Reassigning ``NavigationView.Content`` is unreliable in some win32more bindings,
        # so visibility toggling is used for stable in-place page switching.
        page_host = Grid()
        self._page_host = page_host
        page_host.HorizontalAlignment = HorizontalAlignment.Stretch
        page_host.VerticalAlignment = VerticalAlignment.Stretch
        page_sms.Visibility = Visibility.Visible
        page_orders.Visibility = Visibility.Collapsed
        page_settings.Visibility = Visibility.Collapsed
        page_host.Children.Append(page_sms)
        page_host.Children.Append(page_orders)
        page_host.Children.Append(page_settings)

        nav = NavigationView()
        self._nav = nav
        nav.PaneDisplayMode = NavigationViewPaneDisplayMode.Left
        nav.IsBackButtonVisible = NavigationViewBackButtonVisible.Collapsed
        nav.IsSettingsVisible = False
        nav.IsPaneToggleButtonVisible = False
        nav.OpenPaneLength = 240.0
        nav.PaneTitle = "SafeSales"
        nav.AlwaysShowHeader = True
        nav.VerticalAlignment = VerticalAlignment.Stretch
        nav.HorizontalAlignment = HorizontalAlignment.Stretch

        ni_sms = NavigationViewItem()
        self._nav_item_sms = ni_sms
        ni_sms.Tag = "sms"
        ni_sms.Content = "SMS"
        ni_sms.Icon = SymbolIcon(Symbol.Mail)

        ni_orders = NavigationViewItem()
        self._nav_item_orders = ni_orders
        ni_orders.Tag = "orders"
        ni_orders.Content = "Orders"
        ni_orders.Icon = SymbolIcon(Symbol.Library)

        ni_settings = NavigationViewItem()
        self._nav_item_settings = ni_settings
        ni_settings.Tag = "settings"
        ni_settings.Content = "Settings"
        ni_settings.Icon = SymbolIcon(Symbol.Setting)

        nav.MenuItems.Append(ni_sms)
        nav.MenuItems.Append(ni_orders)
        nav.MenuItems.Append(ni_settings)

        nav.Header = self._build_top_header()
        nav.Content = page_host
        nav.SelectedItem = ni_sms
        nav.SelectionChanged += self._on_nav_selection_changed
        nav.ItemInvoked += self._on_nav_item_invoked

        Grid.SetRow(nav, 0)
        root.Children.Append(nav)

        log_wrap = Border()
        self._log_wrap = log_wrap
        log_wrap.Margin = Thickness(20, 4, 20, 16)
        log_wrap.Padding = Thickness(18, 12, 18, 12)
        log_wrap.BorderThickness = Thickness(1, 1, 1, 1)
        log_wrap.CornerRadius = CornerRadiusHelper.FromUniformRadius(12)
        self._themed_borders.append(log_wrap)

        log_stack = StackPanel()
        log_stack.Orientation = Orientation.Vertical
        log_stack.Spacing = 8

        log_header_row = StackPanel()
        log_header_row.Orientation = Orientation.Horizontal
        log_header_row.Spacing = 8

        log_header = TextBlock()
        self._log_header = log_header
        log_header.Text = "Activity"
        log_header.FontSize = 13
        log_header.FontWeight = FontWeights.SemiBold
        log_header.VerticalAlignment = VerticalAlignment.Center

        log_header_row.Children.Append(log_header)

        self._log_box = TextBox()
        self._log_box.IsReadOnly = True
        self._log_box.AcceptsReturn = True
        self._log_box.TextWrapping = TextWrapping.Wrap
        self._log_box.FontFamily = FontFamily("Cascadia Mono,Consolas")
        self._log_box.FontSize = 12
        self._log_box.BorderThickness = Thickness(0, 0, 0, 0)
        self._log_box.MinHeight = 108

        log_stack.Children.Append(log_header_row)
        log_stack.Children.Append(self._log_box)
        log_wrap.Child = log_stack

        Grid.SetRow(log_wrap, 1)
        root.Children.Append(log_wrap)

        root.ActualThemeChanged += self._on_root_actual_theme_changed

        self._window.Content = root
        self._window.Activate()

        self._setup_pulse_timer()
        self._apply_theme_request(update_combo=True, defer=False)
        p0 = self._active_palette()
        if self._api_dot is not None:
            self._api_dot.Fill = _rgb(255, *p0.offline)
        self._log("Ready.")
        self._refresh_account_async()

    def _on_root_actual_theme_changed(self, sender, args) -> None:
        if self._theme_mode != "system":
            return
        self._sync_chrome_brushes()
        self._set_api_indicator(self._api_online)

    def _effective_dark(self) -> bool:
        if self._theme_mode == "dark":
            return True
        if self._theme_mode == "light":
            return False
        if self._root is None:
            return False
        return self._root.ActualTheme == ElementTheme.Dark

    def _active_palette(self) -> _ChromePalette:
        return _palette_dark() if self._effective_dark() else _palette_light()

    def _sync_chrome_brushes(self) -> None:
        if self._root is None:
            return
        p = self._active_palette()
        self._root.Background = _rgb(255, *p.surface)
        for b in self._themed_borders:
            b.Background = _rgb(255, *p.card)
            b.BorderBrush = _rgb(255, *p.border)
        if self._log_header is not None:
            self._log_header.Foreground = _rgb(255, *p.primary)
        if self._balance_text is not None:
            self._balance_text.Foreground = _rgb(255, *p.primary)
        balance_label = getattr(self, "_balance_label", None)
        if balance_label is not None:
            balance_label.Foreground = _rgb(255, *p.muted)
        if self._api_status_text is not None:
            self._api_status_text.Foreground = _rgb(255, *p.muted)
        account_divider = getattr(self, "_account_divider", None)
        if account_divider is not None:
            account_divider.Background = _rgb(255, *p.border)
        if self._nav_header_title is not None:
            self._nav_header_title.Foreground = _rgb(255, *p.primary)
        if self._nav_header_subtitle is not None:
            self._nav_header_subtitle.Foreground = _rgb(255, *p.muted)
        if self._settings_hint is not None:
            self._settings_hint.Foreground = _rgb(255, *p.muted)
        for ref_name in ("_sms_subtitle", "_orders_subtitle", "_settings_subtitle"):
            tb = getattr(self, ref_name, None)
            if tb is not None:
                tb.Foreground = _rgb(255, *p.muted)
        for ref_name in ("_sms_title", "_orders_title", "_settings_title"):
            tb = getattr(self, ref_name, None)
            if tb is not None:
                tb.Foreground = _rgb(255, *p.primary)
        if self._log_box is not None:
            self._log_box.Foreground = _rgb(255, *p.primary)
            self._log_box.Background = _rgb(255, *p.surface)
        if self._excel_results is not None:
            self._excel_results.Foreground = _rgb(255, *p.primary)
            self._excel_results.Background = _rgb(255, *p.card)
        if self._sms_status is not None:
            self._sms_status.Foreground = _rgb(255, *p.muted)
        if self._send_btn is not None:
            self._send_btn.Background = _rgb(255, *p.accent)
            self._send_btn.Foreground = _rgb(255, *p.on_accent)
        if self._bulk_btn is not None:
            self._bulk_btn.Background = _rgb(255, *p.accent)
            self._bulk_btn.Foreground = _rgb(255, *p.on_accent)

    def _apply_theme_request(self, *, update_combo: bool, defer: bool) -> None:
        def work() -> None:
            self._apply_theme_immediate(update_combo=update_combo)

        if defer:
            self._enqueue(work)
        else:
            work()

    def _apply_theme_immediate(self, *, update_combo: bool) -> None:
        """Apply root element theme and custom chrome. Avoid ``Application.RequestedTheme`` — it can
        invalidate the visual tree during control events (e.g. theme pickers) and crash win32more."""
        if self._root is None:
            return
        if self._theme_mode == "system":
            self._root.RequestedTheme = ElementTheme.Default
        elif self._theme_mode == "light":
            self._root.RequestedTheme = ElementTheme.Light
        else:
            self._root.RequestedTheme = ElementTheme.Dark

        if update_combo and self._theme_combo is not None:
            idx = {"system": 0, "light": 1, "dark": 2}.get(self._theme_mode, 0)
            self._theme_combo_guard = True
            self._theme_combo.SelectedIndex = idx
            self._theme_combo_guard = False

        self._sync_chrome_brushes()

    def _on_theme_combo_selection(self, sender, args) -> None:
        if self._theme_combo_guard or self._theme_combo is None:
            return
        idx = self._theme_combo.SelectedIndex
        modes = ("system", "light", "dark")
        if idx < 0 or idx >= len(modes):
            return
        mode = modes[idx]

        def apply() -> None:
            self._theme_mode = mode
            self._apply_theme_immediate(update_combo=False)
            self._set_api_indicator(self._api_online)

        self._enqueue(apply)

    def _resolve_nav_route(self, item) -> str | None:
        if item is None:
            return None
        route_map = {
            "sms": "sms",
            "one-time sms": "sms",
            "orders": "orders",
            "settings": "settings",
        }
        try:
            tag = item.Tag
        except Exception:  # noqa: BLE001
            tag = None
        if tag is not None:
            route = str(tag).strip().lower()
            mapped = route_map.get(route)
            if mapped is not None:
                return mapped
        try:
            content = item.Content
        except Exception:  # noqa: BLE001
            content = None
        if content is not None:
            route = str(content).strip().lower()
            mapped = route_map.get(route)
            if mapped is not None:
                return mapped
        # In some NavigationView bindings, SelectedItem is the raw content string.
        return route_map.get(str(item).strip().lower())

    def _switch_nav_route(self, route: str, *, sync_selection: bool = True) -> None:
        nav = self._nav
        if nav is None or route not in ("sms", "orders", "settings"):
            return
        self._current_route = route

        pages = {
            "sms": (self._page_sms, self._nav_item_sms, "One-time SMS", "Single recipient delivery"),
            "orders": (self._page_orders, self._nav_item_orders, "Orders", "Excel import and bulk dispatch"),
            "settings": (self._page_settings, self._nav_item_settings, "Settings", "Appearance and application preferences"),
        }
        active_page, active_item, header_title, header_subtitle = pages[route]

        for key, (page, _item, _t, _s) in pages.items():
            if page is None:
                continue
            page.Visibility = Visibility.Visible if key == route else Visibility.Collapsed

        if sync_selection and active_item is not None:
            try:
                nav.SelectedItem = active_item
            except Exception:  # noqa: BLE001 — selection sync should not break navigation
                pass

        if self._nav_header_title is not None:
            self._nav_header_title.Text = header_title
        if self._nav_header_subtitle is not None:
            self._nav_header_subtitle.Text = header_subtitle

    def _on_nav_selection_changed(self, sender, args) -> None:
        item = None
        if sender is not None:
            try:
                item = sender.SelectedItem
            except Exception:  # noqa: BLE001
                item = None
        route = self._resolve_nav_route(item)
        if route is not None:
            self._switch_nav_route(route, sync_selection=False)

    def _on_nav_item_invoked(self, sender, args) -> None:
        try:
            if args.IsSettingsInvoked:
                return
        except Exception:  # noqa: BLE001
            pass
        route: str | None = None
        for accessor in ("InvokedItemContainer", "InvokedItem"):
            try:
                candidate = getattr(args, accessor)
            except Exception:  # noqa: BLE001
                candidate = None
            route = self._resolve_nav_route(candidate)
            if route is not None:
                break
        if route is not None:
            self._switch_nav_route(route, sync_selection=True)

    def _build_top_header(self) -> Grid:
        g = Grid()
        g.VerticalAlignment = VerticalAlignment.Stretch
        g.Margin = Thickness(0, 0, 4, 0)
        c0 = ColumnDefinition()
        c0.Width = GridLengthHelper.FromValueAndType(1.0, GridUnitType.Star)
        c1 = ColumnDefinition()
        c1.Width = GridLengthHelper.FromValueAndType(1.0, GridUnitType.Auto)
        g.ColumnDefinitions.Append(c0)
        g.ColumnDefinitions.Append(c1)

        left = StackPanel()
        left.Orientation = Orientation.Vertical
        left.Spacing = 2
        left.VerticalAlignment = VerticalAlignment.Center
        left.Margin = Thickness(4, 0, 12, 0)
        Grid.SetColumn(left, 0)

        title = TextBlock()
        self._nav_header_title = title
        title.Text = "One-time SMS"
        title.FontSize = 18
        title.FontWeight = FontWeights.SemiBold
        title.VerticalAlignment = VerticalAlignment.Center

        subtitle = TextBlock()
        self._nav_header_subtitle = subtitle
        subtitle.Text = "Single recipient delivery"
        subtitle.FontSize = 12
        subtitle.VerticalAlignment = VerticalAlignment.Center

        left.Children.Append(title)
        left.Children.Append(subtitle)

        account = self._build_account_strip()
        Grid.SetColumn(account, 1)

        g.Children.Append(left)
        g.Children.Append(account)
        return g

    def _setup_pulse_timer(self) -> None:
        t = self._dispatcher.CreateTimer()
        span = TimeSpan()
        span.Duration = 7_500_000
        t.Interval = span
        t.IsRepeating = True
        self._pulse_timer = t
        t.Tick += self._on_pulse_tick

    def _on_pulse_tick(self, sender, args) -> None:
        if not self._api_online or self._api_dot is None:
            return
        self._pulse_phase = not self._pulse_phase
        self._api_dot.Opacity = 1.0 if self._pulse_phase else 0.38

    def _row_star(self) -> RowDefinition:
        rd = RowDefinition()
        rd.Height = GridLengthHelper.FromValueAndType(1.0, GridUnitType.Star)
        return rd

    def _row_px(self, h: float) -> RowDefinition:
        rd = RowDefinition()
        rd.Height = GridLengthHelper.FromPixels(h)
        return rd

    def _register_card(self, b: Border) -> Border:
        self._themed_borders.append(b)
        return b

    def _card(self, inner, *, pad: float = 24.0) -> Border:
        b = Border()
        b.Padding = Thickness(pad, pad, pad, pad)
        b.BorderThickness = Thickness(1, 1, 1, 1)
        b.CornerRadius = CornerRadiusHelper.FromUniformRadius(14)
        b.Child = inner
        b.HorizontalAlignment = HorizontalAlignment.Stretch
        b.MaxWidth = 1400
        return self._register_card(b)

    def _workspace(self, main_card: Border, side_card: Border | None = None) -> Grid:
        g = Grid()
        c0 = ColumnDefinition()
        c0.Width = GridLengthHelper.FromValueAndType(1.0, GridUnitType.Star)
        g.ColumnDefinitions.Append(c0)
        if side_card is not None:
            c1 = ColumnDefinition()
            c1.Width = GridLengthHelper.FromPixels(360)
            g.ColumnDefinitions.Append(c1)
        main_card.Margin = Thickness(20, 0, 12 if side_card is not None else 20, 0)
        Grid.SetColumn(main_card, 0)
        g.Children.Append(main_card)
        if side_card is not None:
            side_card.Margin = Thickness(12, 0, 20, 0)
            Grid.SetColumn(side_card, 1)
            g.Children.Append(side_card)
        return g

    def _page_title(self, text: str, subtitle: str) -> tuple[StackPanel, TextBlock, TextBlock]:
        block = StackPanel()
        block.Orientation = Orientation.Vertical
        block.Spacing = 4
        block.Margin = Thickness(0, 0, 0, 8)

        title = TextBlock()
        title.Text = text
        title.FontSize = 26
        title.FontWeight = FontWeights.SemiBold

        sub = TextBlock()
        sub.Text = subtitle
        sub.FontSize = 13
        sub.TextWrapping = TextWrapping.Wrap

        block.Children.Append(title)
        block.Children.Append(sub)
        return block, title, sub

    def _section_header(self, text: str) -> TextBlock:
        tb = TextBlock()
        tb.Text = text
        tb.FontSize = 15
        tb.FontWeight = FontWeights.SemiBold
        tb.Margin = Thickness(0, 6, 0, 2)
        return tb

    def _build_settings_tab(self) -> ScrollViewer:
        scroll = ScrollViewer()
        scroll.HorizontalScrollMode = ScrollMode.Disabled
        scroll.VerticalScrollBarVisibility = ScrollBarVisibility.Auto
        scroll.Padding = Thickness(24, 24, 24, 32)

        main = StackPanel()
        main.Spacing = 10
        main.Orientation = Orientation.Vertical

        header, title_tb, sub_tb = self._page_title(
            "Settings",
            "Manage how SafeSales SMS looks and behaves.",
        )
        self._settings_title = title_tb
        self._settings_subtitle = sub_tb

        appearance_header = self._section_header("Appearance")

        hint = TextBlock()
        self._settings_hint = hint
        hint.Text = "Match Windows or pin the app to a specific theme."
        hint.TextWrapping = TextWrapping.Wrap
        hint.FontSize = 13

        combo = ComboBox()
        self._theme_combo = combo
        combo.Header = "App theme"
        combo.MinWidth = 280
        for label in ("Use system setting", "Light", "Dark"):
            it = ComboBoxItem()
            it.Content = label
            combo.Items.Append(it)
        self._theme_combo_guard = True
        combo.SelectedIndex = 0
        self._theme_combo_guard = False
        combo.SelectionChanged += self._on_theme_combo_selection

        main.Children.Append(header)
        main.Children.Append(appearance_header)
        main.Children.Append(hint)
        main.Children.Append(combo)

        side = StackPanel()
        side.Spacing = 8
        side.Orientation = Orientation.Vertical

        side_title = TextBlock()
        side_title.Text = "About"
        side_title.FontSize = 15
        side_title.FontWeight = FontWeights.SemiBold

        side_body = TextBlock()
        side_body.Text = (
            "SafeSales desktop client for EasySMS.\n\n"
            "The shell follows Windows 11 Fluent design: Mica backdrop, layered cards, and consistent spacing."
        )
        side_body.TextWrapping = TextWrapping.Wrap
        side_body.FontSize = 13

        side.Children.Append(side_title)
        side.Children.Append(side_body)

        scroll.Content = self._workspace(self._card(main, pad=28), self._card(side, pad=20))
        return scroll

    def _build_account_strip(self) -> Border:
        outer = Border()
        self._account_outer = outer
        outer.Margin = Thickness(0, 0, 0, 0)
        outer.Padding = Thickness(14, 8, 14, 8)
        outer.BorderThickness = Thickness(1, 1, 1, 1)
        outer.CornerRadius = CornerRadiusHelper.FromUniformRadius(10)
        outer.HorizontalAlignment = HorizontalAlignment.Stretch
        outer.VerticalAlignment = VerticalAlignment.Center
        self._register_card(outer)

        row = StackPanel()
        row.Orientation = Orientation.Horizontal
        row.HorizontalAlignment = HorizontalAlignment.Right
        row.Spacing = 12
        row.VerticalAlignment = VerticalAlignment.Center

        balance_label = TextBlock()
        balance_label.Text = "Balance"
        balance_label.FontSize = 12
        balance_label.VerticalAlignment = VerticalAlignment.Center
        self._balance_label = balance_label

        self._balance_text = TextBlock()
        self._balance_text.Text = "—"
        self._balance_text.FontSize = 13
        self._balance_text.FontWeight = FontWeights.SemiBold
        self._balance_text.VerticalAlignment = VerticalAlignment.Center

        refresh = Button()
        refresh.Content = "Refresh"
        refresh.MinWidth = 88
        refresh.VerticalAlignment = VerticalAlignment.Center
        refresh.add_Click(self._on_refresh_balance)

        divider = Border()
        divider.Width = 1
        divider.Height = 18
        divider.Margin = Thickness(2, 0, 2, 0)
        divider.VerticalAlignment = VerticalAlignment.Center
        self._account_divider = divider

        self._api_dot = Ellipse()
        self._api_dot.Width = 9
        self._api_dot.Height = 9
        self._api_dot.VerticalAlignment = VerticalAlignment.Center
        self._api_dot.Opacity = 1.0

        self._api_status_text = TextBlock()
        self._api_status_text.Text = "Checking…"
        self._api_status_text.FontSize = 12
        self._api_status_text.VerticalAlignment = VerticalAlignment.Center

        status_row = StackPanel()
        status_row.Orientation = Orientation.Horizontal
        status_row.Spacing = 6
        status_row.VerticalAlignment = VerticalAlignment.Center
        status_row.Children.Append(self._api_dot)
        status_row.Children.Append(self._api_status_text)

        row.Children.Append(balance_label)
        row.Children.Append(self._balance_text)
        row.Children.Append(refresh)
        row.Children.Append(divider)
        row.Children.Append(status_row)

        outer.Child = row
        return outer

    def _centered_tab(self, inner: StackPanel) -> Grid:
        g = Grid()
        inner.HorizontalAlignment = HorizontalAlignment.Stretch
        inner.VerticalAlignment = VerticalAlignment.Top
        inner.MaxWidth = 1040
        card = self._card(inner, pad=24)
        card.HorizontalAlignment = HorizontalAlignment.Stretch
        card.Margin = Thickness(20, 0, 20, 0)
        g.Children.Append(card)
        return g

    def _build_one_time_tab(self) -> ScrollViewer:
        scroll = ScrollViewer()
        scroll.HorizontalScrollMode = ScrollMode.Disabled
        scroll.VerticalScrollBarVisibility = ScrollBarVisibility.Auto
        scroll.Padding = Thickness(24, 24, 24, 32)

        panel = StackPanel()
        panel.Spacing = 14
        panel.Orientation = Orientation.Vertical

        header, title_tb, sub_tb = self._page_title(
            "Compose message",
            "Send a one-time SMS with optional sender ID.",
        )
        self._sms_title = title_tb
        self._sms_subtitle = sub_tb

        self._sms_phone = TextBox()
        self._sms_phone.Header = "To"
        self._sms_phone.PlaceholderText = "Mobile number"

        self._sms_message = TextBox()
        self._sms_message.Header = "Message"
        self._sms_message.AcceptsReturn = True
        self._sms_message.TextWrapping = TextWrapping.Wrap
        self._sms_message.MinHeight = 128

        self._sms_sender = TextBox()
        self._sms_sender.Header = "Sender (optional)"
        self._sms_sender.PlaceholderText = "Sender ID"

        send_btn = Button()
        self._send_btn = send_btn
        send_btn.Content = "Send message"
        send_btn.MinWidth = 168
        send_btn.HorizontalAlignment = HorizontalAlignment.Left
        send_btn.Margin = Thickness(0, 4, 0, 0)
        send_btn.add_Click(self._on_send_one)

        self._sms_status = TextBlock()
        self._sms_status.Text = ""
        self._sms_status.TextWrapping = TextWrapping.Wrap
        self._sms_status.FontSize = 12
        self._sms_status.Margin = Thickness(0, 2, 0, 0)

        panel.Children.Append(header)
        panel.Children.Append(self._sms_phone)
        panel.Children.Append(self._sms_message)
        panel.Children.Append(self._sms_sender)
        panel.Children.Append(send_btn)
        panel.Children.Append(self._sms_status)

        side = StackPanel()
        side.Spacing = 8
        side.Orientation = Orientation.Vertical

        side_title = TextBlock()
        side_title.Text = "Tips"
        side_title.FontSize = 15
        side_title.FontWeight = FontWeights.SemiBold

        side_body = TextBlock()
        side_body.Text = (
            "• Use international number format when possible.\n"
            "• Keep sender short for better carrier compatibility.\n"
            "• Review the message before sending — SMS cannot be undone."
        )
        side_body.TextWrapping = TextWrapping.Wrap
        side_body.FontSize = 13
        side_body.LineHeight = 20

        side.Children.Append(side_title)
        side.Children.Append(side_body)

        scroll.Content = self._workspace(self._card(panel, pad=28), self._card(side, pad=20))
        return scroll

    def _file_row(self, text_box: TextBox, button: Button) -> Grid:
        g = Grid()
        c0 = ColumnDefinition()
        c0.Width = GridLengthHelper.FromValueAndType(1.0, GridUnitType.Star)
        c1 = ColumnDefinition()
        c1.Width = GridLengthHelper.FromValueAndType(1.0, GridUnitType.Auto)
        g.ColumnDefinitions.Append(c0)
        g.ColumnDefinitions.Append(c1)
        text_box.Margin = Thickness(0, 0, 8, 0)
        Grid.SetColumn(text_box, 0)
        button.VerticalAlignment = VerticalAlignment.Bottom
        Grid.SetColumn(button, 1)
        g.Children.Append(text_box)
        g.Children.Append(button)
        return g

    def _build_excel_tab(self) -> ScrollViewer:
        scroll = ScrollViewer()
        scroll.HorizontalScrollMode = ScrollMode.Disabled
        scroll.VerticalScrollBarVisibility = ScrollBarVisibility.Auto
        scroll.Padding = Thickness(24, 24, 24, 32)

        panel = StackPanel()
        panel.Spacing = 12
        panel.Orientation = Orientation.Vertical

        header, title_tb, sub_tb = self._page_title(
            "Orders",
            "Import ACS / Box Express sheets, validate rows, then send in bulk.",
        )
        self._orders_title = title_tb
        self._orders_subtitle = sub_tb

        files_header = self._section_header("Source files")

        self._acs_path_box = TextBox()
        self._acs_path_box.Header = "ACS workbook"
        self._acs_path_box.PlaceholderText = "Select an .xlsx file"
        self._acs_path_box.IsReadOnly = True

        acs_btn = Button()
        acs_btn.Content = "Browse"
        acs_btn.MinWidth = 112
        acs_btn.add_Click(self._on_browse_acs)

        self._box_path_box = TextBox()
        self._box_path_box.Header = "Box Express workbook"
        self._box_path_box.PlaceholderText = "Select an .xlsx file"
        self._box_path_box.IsReadOnly = True

        box_btn = Button()
        box_btn.Content = "Browse"
        box_btn.MinWidth = 112
        box_btn.add_Click(self._on_browse_box)

        process_btn = Button()
        process_btn.Content = "Process files"
        process_btn.MinWidth = 168
        process_btn.HorizontalAlignment = HorizontalAlignment.Left
        process_btn.Margin = Thickness(0, 6, 0, 0)
        process_btn.add_Click(self._on_process_excel)

        compose_header = self._section_header("Compose")

        self._excel_template = TextBox()
        self._excel_template.Header = "Message template — {voucher}, {phone}, {carrier}, {source_file}"
        self._excel_template.PlaceholderText = "Hello {phone}, your shipment {voucher} is on the way."
        self._excel_template.TextWrapping = TextWrapping.Wrap
        self._excel_template.AcceptsReturn = True
        self._excel_template.MinHeight = 84

        self._excel_sender = TextBox()
        self._excel_sender.Header = "Sender (optional)"
        self._excel_sender.PlaceholderText = "Sender ID"

        bulk_btn = Button()
        self._bulk_btn = bulk_btn
        bulk_btn.Content = "Send to all rows"
        bulk_btn.MinWidth = 168
        bulk_btn.HorizontalAlignment = HorizontalAlignment.Left
        bulk_btn.Margin = Thickness(0, 4, 0, 0)
        bulk_btn.add_Click(self._on_bulk_send)

        self._progress = ProgressRing()
        self._progress.Width = 24
        self._progress.Height = 24
        self._progress.IsActive = False
        self._progress.Visibility = Visibility.Collapsed

        self._excel_results = TextBox()
        self._excel_results.Header = "Results"
        self._excel_results.IsReadOnly = True
        self._excel_results.TextWrapping = TextWrapping.Wrap
        self._excel_results.AcceptsReturn = True
        self._excel_results.FontFamily = FontFamily("Cascadia Mono,Consolas")
        self._excel_results.FontSize = 12
        self._excel_results.MinHeight = 320

        panel.Children.Append(header)
        panel.Children.Append(files_header)
        panel.Children.Append(self._file_row(self._acs_path_box, acs_btn))
        panel.Children.Append(self._file_row(self._box_path_box, box_btn))
        panel.Children.Append(process_btn)
        panel.Children.Append(compose_header)
        panel.Children.Append(self._excel_template)
        panel.Children.Append(self._excel_sender)
        panel.Children.Append(bulk_btn)

        side = StackPanel()
        side.Spacing = 10
        side.Orientation = Orientation.Vertical

        side_title = TextBlock()
        side_title.Text = "Status"
        side_title.FontSize = 15
        side_title.FontWeight = FontWeights.SemiBold

        status_row = StackPanel()
        status_row.Orientation = Orientation.Horizontal
        status_row.Spacing = 10
        status_row.VerticalAlignment = VerticalAlignment.Center

        progress_label = TextBlock()
        progress_label.Text = "Idle"
        progress_label.FontSize = 12
        progress_label.VerticalAlignment = VerticalAlignment.Center
        self._orders_status_label = progress_label

        status_row.Children.Append(self._progress)
        status_row.Children.Append(progress_label)

        side.Children.Append(side_title)
        side.Children.Append(status_row)
        side.Children.Append(self._excel_results)

        scroll.Content = self._workspace(self._card(panel, pad=28), self._card(side, pad=20))
        return scroll

    def _enqueue(self, fn, *, high: bool = False) -> None:
        prio = DispatcherQueuePriority.High if high else DispatcherQueuePriority.Normal
        self._dispatcher.TryEnqueueWithPriority(prio, fn)

    def _log(self, message: str) -> None:
        line = f"[{_ts()}] {message}\n"

        def append():
            if self._log_box is None:
                return
            cur = self._log_box.Text or ""
            self._log_box.Text = (cur + line)[-120000:]

        if self._dispatcher.HasThreadAccess:
            append()
        else:
            self._enqueue(append)

    def _set_api_indicator(self, online: bool) -> None:
        self._api_online = online
        if self._api_dot is None or self._api_status_text is None:
            return
        p = self._active_palette()
        if online:
            self._api_dot.Fill = _rgb(255, *p.online)
            self._api_status_text.Text = "API online"
            self._api_dot.Opacity = 1.0
            self._pulse_phase = True
            if self._pulse_timer is not None and not self._pulse_timer.IsRunning:
                self._pulse_timer.Start()
        else:
            self._api_dot.Fill = _rgb(255, *p.offline)
            self._api_status_text.Text = "Offline"
            self._api_dot.Opacity = 1.0
            if self._pulse_timer is not None and self._pulse_timer.IsRunning:
                self._pulse_timer.Stop()

    def _apply_account_ui(self, balance_label: str, online: bool) -> None:
        if self._balance_text is not None:
            self._balance_text.Text = balance_label or "—"
        self._set_api_indicator(online)

    def _refresh_account_async(self) -> None:
        client = self._client

        def worker():
            online = services.is_easysms_online(client)
            if client:
                try:
                    bal = services.fetch_balance_display(client)
                except Exception:  # noqa: BLE001
                    bal = "—"
            else:
                bal = "—"
            self._enqueue(lambda: self._apply_account_ui(bal, online))

        threading.Thread(target=worker, daemon=True).start()

    def _on_refresh_balance(self, sender, args):
        self._log("Account refresh.")
        self._refresh_account_async()

    def _on_send_one(self, sender, args):
        client = self._client
        if not client:
            self._sms_status.Text = "Unavailable."
            return
        phone = (self._sms_phone.Text or "").strip()
        text = (self._sms_message.Text or "").strip()
        sender_id = (self._sms_sender.Text or "").strip() or None
        if not phone or not text:
            self._sms_status.Text = "Number and message required."
            return

        self._sms_status.Text = "Sending…"

        def worker():
            try:
                services.send_one_sms(client, phone, text, sender=sender_id)
                self._enqueue(lambda: self._finish_send_ok(f"Sent to {phone}."))
            except EasySMSError as exc:
                self._enqueue(lambda: self._finish_send_err(str(exc)))
            except OSError as exc:
                self._enqueue(lambda: self._finish_send_err(str(exc)))
            except Exception as exc:  # noqa: BLE001
                self._enqueue(lambda: self._finish_send_err(str(exc)))

        threading.Thread(target=worker, daemon=True).start()

    def _finish_send_ok(self, msg: str) -> None:
        self._sms_status.Text = msg
        self._log(msg)
        self._refresh_account_async()

    def _finish_send_err(self, msg: str) -> None:
        self._sms_status.Text = msg
        self._log(f"Error: {msg}")

    def _on_browse_acs(self, sender, args):
        self._pick_target = "acs"
        self._launch_file_picker()

    def _on_browse_box(self, sender, args):
        self._pick_target = "box"
        self._launch_file_picker()

    def _launch_file_picker(self) -> None:
        wid = self._window.AppWindow.Id
        picker = FileOpenPicker(wid)
        picker.SuggestedStartLocation = PickerLocationId.DocumentsLibrary
        picker.FileTypeFilter.Clear()
        picker.FileTypeFilter.Append(".xlsx")

        op = picker.PickSingleFileAsync()

        def completed(ai, st):
            if st != AsyncStatus.Completed:
                return
            try:
                res: PickFileResult | None = ai.GetResults()
            except Exception:
                return
            path = ""
            if res is not None:
                try:
                    path = str(res.Path or "")
                except Exception:
                    return
            target = self._pick_target

            def apply():
                if target == "acs":
                    self._acs_path = path
                    if self._acs_path_box:
                        self._acs_path_box.Text = path
                elif target == "box":
                    self._box_path = path
                    if self._box_path_box:
                        self._box_path_box.Text = path
                if path:
                    self._log(path)
                self._pick_target = None

            self._enqueue(apply)

        op.Completed = completed

    def _set_progress(self, active: bool) -> None:
        if self._progress is not None:
            self._progress.IsActive = active
            self._progress.Visibility = Visibility.Visible if active else Visibility.Collapsed
        label = getattr(self, "_orders_status_label", None)
        if label is not None:
            label.Text = "Working…" if active else "Idle"

    def _on_process_excel(self, sender, args):
        paths: list[Path] = []
        if self._acs_path:
            paths.append(Path(self._acs_path))
        if self._box_path:
            paths.append(Path(self._box_path))
        if not paths:
            self._log("No files selected.")
            return

        self._log("Processing…")
        self._set_progress(True)

        def worker():
            try:
                result = services.process_excel_paths(paths)
            except Exception as exc:  # noqa: BLE001
                tb = traceback.format_exc()
                self._enqueue(lambda: self._finish_process(None, f"{exc}\n{tb}", f"{exc}\n{tb}"))
                return

            panel = services.format_excel_results_panel(result.lines, result.shipments)
            lines = "\n".join(result.lines)
            self._enqueue(lambda: self._finish_process(result.shipments, panel, lines))

        threading.Thread(target=worker, daemon=True).start()

    def _finish_process(self, shipments, results_text: str | None, log_block: str) -> None:
        self._set_progress(False)
        if shipments is not None:
            self._last_shipments = shipments
        if self._excel_results is not None and results_text is not None:
            self._excel_results.Text = results_text
        if shipments is not None:
            self._log("Done.")
        elif results_text:
            self._log("Process failed.")
        for line in (log_block or "").splitlines():
            if line.strip():
                self._log(line)

    def _on_bulk_send(self, sender, args):
        client = self._client
        if not client:
            self._log("Unavailable.")
            return
        if not self._last_shipments:
            self._log("Process files first.")
            return
        template = (self._excel_template.Text or "").strip()
        if not template:
            self._log("Message template empty.")
            return

        sender_id = (self._excel_sender.Text or "").strip() or None
        self._log("Sending…")
        self._set_progress(True)

        shipments_snapshot = list(self._last_shipments)

        def worker():
            try:
                lines = services.send_shipments(client, shipments_snapshot, template, sender=sender_id)
            except Exception as exc:  # noqa: BLE001
                err_tb = traceback.format_exc()
                self._enqueue(lambda: self._finish_bulk(None, f"{exc}\n{err_tb}"))
                return
            self._enqueue(lambda lns=lines: self._finish_bulk(lns, None))

        threading.Thread(target=worker, daemon=True).start()

    def _finish_bulk(self, lines: list[str] | None, fatal: str | None) -> None:
        self._set_progress(False)
        if fatal:
            self._log(fatal)
            self._refresh_account_async()
            return
        if not lines:
            self._refresh_account_async()
            return
        for line in lines or []:
            self._log(line)
        self._refresh_account_async()


def main() -> None:
    XamlApplication.Start(SafeSalesWinUIApp)


if __name__ == "__main__":
    main()
