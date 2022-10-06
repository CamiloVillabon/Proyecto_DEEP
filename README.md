# Vibrava
Vibration Accumulated Report Automation for Downhole Tools

Vibrava
============================================================

Implementation of Accumulated Report Automation for Downhole Tools using Python, Tkinter and Power BI. Particularly, the pipeline provides support for:

- Sperry Drilling
- Landmark
- Drill Bits Services

Vibrava is responsible for acummulating vibration data with respect to every downhole tool, creating a full diagnosis in order to detect vibration mechanisms, decreasing operation time and costs. Then, this data will be displayed by using a Power BI dashboard, Final Excel Report and User Interface.  

When the calculation of the delay per provider is finished, these metrics are sent to Cloudwatch in a time interval of 5 minutes.

# Getting Started

You can install the most recent version of the pipeline with

```zsh
pip install git+:https://github.com/CamiloVillabon/Proyecto_DEEP.git
```

# To test in local

Advice: create a new conda env with version of python 3.10.4

Install requirements.txt of the galactus-metrics and the whole replication repo
```
vibrava$ pip install -r requiremets.txt
```

run as a administrator to ./vibrava.exe
remember to be connected to the Halliburton domain.
```

# Next Steps

    - Database connection (Insite, SAP, etc).
    - Improving data visualization
