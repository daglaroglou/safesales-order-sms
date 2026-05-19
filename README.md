<p align="center">
  <img src="https://capsule-render.vercel.app/api?type=waving&height=220&text=SafeSales%20Order%20SMS&fontAlign=50&fontAlignY=38&color=0:0f172a,50:1d4ed8,100:38bdf8&fontColor=ffffff&desc=Smart%20desktop%20SMS%20dispatch%20for%20daily%20shipments&descAlign=50&descAlignY=58" alt="SafeSales banner" />
</p>

<p align="center">
  <img src="https://readme-typing-svg.demolab.com?font=Segoe+UI&weight=600&size=22&duration=2300&pause=700&color=1D4ED8&center=true&vCenter=true&width=900&lines=One-time+and+bulk+SMS+sending;ACS+%2B+Box+Express+Excel+flow;Courier-aware+tracking+templates;Built+for+fast+daily+operations" alt="Typing animation" />
</p>

<p align="center">
  <img src="https://img.shields.io/badge/Platform-Windows-0078D6?style=for-the-badge&logo=windows&logoColor=white" alt="Windows" />
  <img src="https://img.shields.io/badge/Python-Desktop_App-3776AB?style=for-the-badge&logo=python&logoColor=white" alt="Python" />
  <img src="https://img.shields.io/badge/UI-WinUI_3-0F6CBD?style=for-the-badge" alt="WinUI 3" />
  <img src="https://img.shields.io/badge/SMS-EasySMS-16A34A?style=for-the-badge" alt="EasySMS" />
</p>

---

## What It Is

**SafeSales Order SMS** is a desktop app that helps operations teams send shipment notifications with speed and consistency.

From one clean interface, you can import courier workbooks, validate rows, send one-time or bulk tracking messages, and monitor status in real time.

## Highlights

- Send **one-time SMS** with optional sender ID.
- Run **bulk dispatch from Excel** for ACS and Box Express.
- Use **courier-aware templates** with automatic voucher insertion.
- Maintain **daily courier groups** for contact organization.
- Track **API connectivity and SMS balance** in-app.
- Keep **API key and export settings** persisted for daily use.

## Daily Workflow

1. Add your EasySMS API key in the Settings page.
2. Select ACS and/or Box Express Excel files.
3. Process files and review parsed shipment rows.
4. Choose courier target and sender ID.
5. Send in bulk and follow activity logs.

## Built For

- E-commerce and operations teams
- Back-office shipment processing
- Businesses that need a lightweight Windows SMS tool

## Tech Snapshot

- Python desktop application
- WinUI 3 shell (`win32more`)
- EasySMS API integration
- Excel parsing pipeline for courier sheets

## Installer

A Windows installer is now supported for packaging the built EXE.

To build the installer locally:

1. Install NSIS on the build machine.
2. Run `powerShell -File packaging\build_installer.ps1`.
3. The installer will be produced in `release\`, for example `release\SafeSalesSMSSetup-1.0.0.exe`.

> Note: The Windows App Runtime (Windows App SDK Runtime) is still required on target machines. If the app fails to launch after installation, install the runtime from:
> `https://learn.microsoft.com/windows/apps/windows-app-sdk/prepare-systems`
