# Vertiv Liebert GXT MT+ 3000VA UPS Reader (BMGUPS)

A C# console application for monitoring the **Vertiv Liebert GXT MT+ 3000VA 2400W** over USB.  
Provides real-time power information, logging, and total energy usage tracking.

---

## 🔧 Features

- Auto-detects USB communication settings on first launch
- Displays real-time UPS data in the console (updated every second)
- Optional CSV logging of all output data
- Tracks and stores total energy usage (kWh) in `KWh.txt`
- Lightweight and portable `.exe` — no install required

---

## 📦 Download

👉 [Download BMGUPS.exe](https://github.com/bmgjet/UPS/raw/refs/heads/master/BMGUPS/bin/Release/BMGUPS.exe)

---

## 🚀 Getting Started

### 1. Run the app

Simply start the executable:

```bash
BMGUPS.exe
```
![Screenshot](https://github.com/bmgjet/UPS/blob/master/screenshot.png?raw=true)

On first launch, it will auto-detect the USB settings and generate a configuration file:

```
BMGUPS.config
```

### 2. Configuration file layout

```
UPSSize | Batterys | CostKWh | PowerFactor
```

- `UPSSize`: VA rating of the UPS
- `Batterys`: Number of battery packs
- `CostKWh`: Electricity cost per kWh
- `PowerFactor`: Efficiency (e.g., 0.8)

### 3. CSV Logging (Optional)

To record output to a CSV file:

```bash
BMGUPS.exe mylog.csv
```

This will continuously append new UPS data to `mylog.csv`.

### 4. Power Usage Tracking

Every 5 minutes, `KWh.txt` is updated with cumulative power usage.  
**IMPORTANT:** Always close the app with `CTRL+C` so the final data is saved correctly.

---

## 📂 Files Generated

- `BMGUPS.config` – Configuration file
- `KWh.txt` – Total cumulative energy used
- `[yourfilename].csv` – Optional data log file

---

## 💡 Notes

- Ensure the UPS is connected via USB and powered on.
- Designed specifically for the **Vertiv Liebert GXT MT+ 3000VA** model.
- Requires no additional drivers on Windows for basic USB communication.

---

## 📜 License

This project is released under the MIT License.
