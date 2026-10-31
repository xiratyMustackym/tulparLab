Tulpar Lab



Tulpar Lab is a Python-based desktop engineering suite and simulation environment designed for materials science and electrical engineering analysis. It provides interactive, real-time simulated testing modules, data visualization, and comparative analysis tools for experimental and theoretical material properties.



Features



The application currently features three primary analytical modules:



Tensile Test Analyzer: 
Simulates Force-Extension and Stress-Strain curves (including elastic, yield point, strain hardening, and necking regions). Calculates Modulus of Toughness, Resilience, and critical mechanical properties for standard and low-carbon steels.

4-Point Probe \& I-V Characterization: 
Simulates electrical measurements for thin films and bulk materials. Features dual modes for Ohmic (Linear) and Semiconductor (Non-Linear) behaviors, calculating sheet resistance, bulk resistivity, and thermal voltage responses.

X-Ray Diffraction (XRD) Crystallography: 
Simulates theoretical XRD patterns based on Bragg's Law. Supports multiple crystal systems (Cubic, Tetragonal, Orthorhombic, Hexagonal) and Bravais lattices, generating relative intensity peaks and precise $2\\theta$ angles.

Key Technical Highlights

Real-Time Animation:
Live plotting of test results using `matplotlib` to mimic physical testing environments.

Data Export \& Comparison:
Automatically saves raw test data and summary reports in `.csv` format. Includes a built-in comparison tool that overlays graphs from different materials.

Dynamic Database:
Built-in material database using `.xlsx` files, allowing users to add, delete, and load custom external databases seamlessly.



## Installation & Setup

1. **Clone the repository:**
   ```bash
   git clone https://github.com/xiratyMustackyn/tulparLab.git
   cd tulparLab
