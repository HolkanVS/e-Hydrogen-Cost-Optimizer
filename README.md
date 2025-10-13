<div align="center">

<img src="images/logo_e_h2_squared.svg" alt="e-Hydrogen Cost Optimizat Logo" width="350"/>
<h1>e-Hydrogen Cost Optimizer</h1>
<h3>Python-based User-defined Techno-economic Optimization and Life Cycle Assessment for e-Hydrogen Production </h3>

Concept DOI (all versions) : 
  
[![DOI](https://zenodo.org/badge/DOI/10.5281/zenodo.17198884.svg)](https://doi.org/10.5281/zenodo.17198884)

</div>

---
<div align="center">
  <a href="https://e-h2.org" target="_blank" rel="noopener noreferrer" style="
    display: inline-block;
    background-color: #0f8a33;
    color: white;
    padding: 10px 18px;
    border-radius: 6px;
    text-decoration: none;
    font-weight: 600;
  ">Official website : e-h2.org</a>
</div>

---

## 📑Overview
The **e-Hydrogen Cost Optimizer** application integrates techno-economic optimization and life cycle assessment (LCA) for the production of electrolytic hydrogen (e-hydrogen) powered by renewable energy sources.

The **e-Hydrogen Cost Optimizer** app is built on top of the [Brightway LCA](https://docs.brightway.dev/en/latest/) framework as well as the [optimization modeling framework Pyomo](https://www.pyomo.org/). 

---

## ✏Capabilities
 Its capabilities include:
 ### 🔧 Techno-Economic Modeling and Optimization
- **Mixed-Integer Linear Programming (MILP)** for cost-optimal system design, based on the lowest **Levelized Cost of Hydrogen (LCOH)**.
- Optimization of energy systems comprising:
  - Solar photovoltaics (PV)
  - Wind turbines
  - Battery storage systems
  - Electrolyzers
  - Hydrogen storage tanks
- Customizable project parameters, such as hydrogen demand (daily or yearly), location coordinates, component capital and operational expenditure, system lifetime, etc.
- Scenario-based optimization with selectable technology options for each component.
- Graphical interface to run optimizations and view solver status in real time.

### 🌱 Life Cycle Assessment (LCA)
- Built-in integration with the [Brightway2](https://docs.brightway.dev/en/latest/) LCA framework.
- Calculation of climate impact using standard LCIA methods (e.g., Global Warming Potential).
- Component-level environmental performance analysis.
- Exportable CSV file with detailed LCA results per subcomponent.

### 📉 Output and Results
- Export of comprehensive results to `.xlsx` and `.csv` files for further analysis.
- Includes optimized capacities, energy flows, hydrogen production, storage states,  cost metric, environmental performance per subcomponent, etc.
- Results are organized for easy interpretation and reuse.

---

## 📘 User Manual
A complete **User Manual** is included in the [wiki](https://github.com/HolkanVS/e-Hydrogen-Cost-Optimizer-v0.3/wiki) section to guide you through each feature of the application.

It provides detailed, step-by-step instructions for:

- Setting system parameters and selecting technologies
- Running techno-economic optimization
- Interpreting cost results and decision variables
- Exploring time series of system operation
- Performing environmental impact assessments using LCA
- Exporting and analyzing results
- Replicating real-world case studies

**We highly recommend reviewing the manual before starting a new project.** It is especially useful for first-time users and those working with custom scenarios or advanced inputs.

---

## 💿 How To Install
**Note:** This section provides a quick overview for installing the application.  
> For the complete installation procedure, please refer to the full [Installation](https://github.com/HolkanVS/e-Hydrogen-Cost-Optimizer-v0.3/wiki/Installation) guide in the wiki section.

Follow these steps to install and run the **e-Hydrogen Cost Optimizer** on a Windows system:

### System Requirements
- Operating System: **Windows 10 or higher**
- Internet connection (required for downloading weather data and LCA background data)
- No Python installation required — the app runs as a standalone executable (`.exe`) alongside with an `_internal` folder that contains all the resources required.

### 📥 Installation Steps

1. **Download the Application**
   - Get the latest version from the official source:

     [Download e_Hydrogen_Cost_Optimizer_v_0_3_1.zip](https://kaust-my.sharepoint.com/:u:/g/personal/vazqueh_kaust_edu_sa/EcV8zr_Tum1MrNZYRoj2_K8Bh4IKeMUboWhOfY0hlEhFdA?e=ODEOMl)

2. **Extract the Archive**
   - Right-click on the downloaded `.zip` file and select **“Extract All...”**
   - This will create a folder named `e_Hydrogen_Cost_Optimizer_v_0_3_1`

3. **Run the Application**
   - Open the extracted folder and double-click the file:
          `e_Hydrogen_Cost_Optimizer_v_0_3_1.exe`

    - **⚠ Important:** Do **not** move the `.exe` file out of its folder. If you need to relocate the executable, move the entire folder to a new location.
   - You may optionally create a shortcut to the `.exe` file on your desktop.

4. **Uninstallation**
   - To uninstall the application, simply delete the folder `e_Hydrogen_Cost_Optimizer_v_0_3_1`.

---

## ⌨ Source Code 
The source code is mainly found at the [hydrogen_optimizer_v_0_3_1.py](hydrogen_optimizer_v_0_3_1.py) `python` file. 

---

## ✍Authors
The **e-Hydrogen Cost Optimizer** is being continuously developed and maintained by researchers at  
**King Abdullah University of Science and Technology (KAUST)**.

This tool is the result of ongoing collaboration between experts in renewable energy systems, optimization modeling, and environmental assessment.

For academic inquiries, collaborations, or feature requests, please contact the development team:

- **Holkan Vazquez-Sanchez**  
  📧 [holkan.vazquezsanchez@kaust.edu.sa](mailto:holkan.vazquezsanchez@kaust.edu.sa)

- **Monserrat Echegoyen Lopez**    
  📧[monserrat.lopez@kaust.edu.sa](mailto:monserrat.lopez@kaust.edu.sa)

- **Dr. Mani Sarathy**  
  📧 [mani.sarathy@kaust.edu.sa ](mailto:mani.sarathy@kaust.edu.sa )

- **Dr. Aziz Nechache**   
  📧 [aziz.nechache@kaust.edu.sa](mailto:aziz.nechache@kaust.edu.sa )

--- 
## :clipboard: Citation


*This project is actively evolving — contributions and feedback are welcome!*
