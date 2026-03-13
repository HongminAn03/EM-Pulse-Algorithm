# EM Dimming Algorithm

## Overview
This project implements an **Electro-Magnetic (EM) dimming algorithm** to reduce flicker in display panels that use **Pulse-Width Modulation (PWM)**.  
The algorithm ensures more stable luminance across varying brightness levels by minimizing uneven off-times within a frame.

## Features
- Recursive mapping of odd/even pulses to **distribute off-times evenly**  
- Flicker evaluation tool that calculates **standard deviation of off-times** for stability measurement  
- Excel-based reporting with **frame-by-frame pulse visualization** for easy validation  

## Tech Stack
- **Language:** Python  
- **Libraries:** `openpyxl`, `matplotlib`, `numpy`  

## Getting Started

### Prerequisites
- Python 3.7+

### Installation
```bash
git clone https://github.com/HongminAn03/EM-Pulse-Algorithm.git
cd EM-Pulse-Algorithm
pip install -r requirements.txt
```

### Usage
```bash
python Gen_EMPS.py
```

The script runs interactively. When prompted, enter the number of pulses (e.g. `32`). It will:
1. Generate the deterministic EM pulse sequence
2. Print the sequence and flicker metrics to the terminal
3. Save a pulse sequence diagram as `em_pulse_seq_<n>.xlsx`

Type `exit` to quit.

## Example Output
- **Excel export** of pulse sequences (`em_pulse_seq_32.xlsx`)
- **Flicker metrics** (max and average standard deviation of off-times) printed to console

## Use Cases
- Display Driver IC algorithm development  
- Flicker evaluation in PWM-driven displays  
- Hardware-friendly iterative approaches to pulse scheduling  

---