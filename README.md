<div align="center">

<img src="images/logo_e_h2_squared.svg" alt="e-Hydrogen Cost Optimizat Logo" width="350"/>
<h1>e-Hydrogen Cost Optimizer</h1>
<h3>Python-based User-defined Techno-economic Optimization and Life Cycle Assessment for e-Hydrogen Production </h3>

  
[![DOI](https://zenodo.org/badge/DOI/10.5281/zenodo.17198884.svg)](https://doi.org/10.5281/zenodo.17198884)
[![Static Badge](https://img.shields.io/badge/official_website-e--h2.org-%230f8a33)](https://e-h2.org)
![GitHub file size in bytes](https://img.shields.io/github/size/HolkanVS/e-Hydrogen-Cost-Optimizer/docs%2Fdownloads%2Fe-Hydrogen%20Cost%20Optimizer%20v1.0.0.zip?label=application%20size(.zip)&color=%235C8C46)


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
- Built-in integration with the [Brightway25](https://docs.brightway.dev/en/latest/) LCA framework.
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
  
> For the complete installation procedure, please refer to the full [Installation](https://github.com/HolkanVS/e-Hydrogen-Cost-Optimizer-v0.3/wiki/Installation) guide in the wiki section.


---

## ⌨ Source Code 
The source code is mainly found at the [hydrogen_optimizer_v_1_0_0.py](hydrogen_optimizer_v_1_0_0.py) `python` file. 

---

## ✍Authors
The **e-Hydrogen Cost Optimizer** is being continuously developed and maintained by researchers at  
**King Abdullah University of Science and Technology (KAUST)**.

This tool is the result of ongoing collaboration between experts in renewable energy systems, optimization modeling, and environmental assessment.

For academic inquiries, collaborations, or feature requests, please contact the development team:

- **Holkan Vazquez-Sanchez**  
  📧 [holkan.vazquezsanchez@kaust.edu.sa](mailto:holkan.vazquezsanchez@kaust.edu.sa)

- **Chengcheng Zhao**  
  📧[chengcheng.zhao@kaust.edu.sa](mailto:chengcheng.zhao@kaust.edu.sa)

- **Monserrat Echegoyen-Lopez**    
  📧[monserrat.lopez@kaust.edu.sa](mailto:monserrat.lopez@kaust.edu.sa)

- **Dr. Mani Sarathy**  
  📧 [mani.sarathy@kaust.edu.sa ](mailto:mani.sarathy@kaust.edu.sa )

- **Dr. Aziz Nechache**   
  📧 [aziz.nechache@kaust.edu.sa](mailto:aziz.nechache@kaust.edu.sa )

--- 
## :clipboard: Citation
### APA
Vazquez-Sanchez, H. (2025). e-Hydrogen Cost Optimizer (Version : latest) [Computer software]. https://doi.org/10.5281/zenodo.17198884

### BibText
@software{Vazquez-Sanchez_e-Hydrogen_Cost_Optimizer_2025,  
author = {Vazquez-Sanchez, Holkan},  
doi = {10.5281/zenodo.17198884},  
license = {Apache-2.0},  
month = oct,  
title = {{e-Hydrogen Cost Optimizer}},  
url = {https://github.com/HolkanVS/e-Hydrogen-Cost-Optimizer/},  
version = {: latest},  
year = {2025}  
}

---

*This project is actively evolving — contributions and feedback are welcome!*
