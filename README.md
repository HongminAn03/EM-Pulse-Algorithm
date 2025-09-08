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

## Example Output
- **Excel export** of pulse sequences  
- **Visualization graphs** showing pulse timing per frame  

## Use Cases
- Display Driver IC algorithm development  
- Flicker evaluation in PWM-driven displays  
- Hardware-friendly iterative approaches to pulse scheduling  

---
