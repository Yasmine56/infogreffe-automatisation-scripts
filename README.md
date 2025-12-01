# Automated Statistics Generation for Infogreffe

This repository contains Python scripts to automate the generation of statistics for different regions of France and prepare a PowerPoint report with updated slides that I developed during my internship at Infogreffe.

## Scripts Overview

- `national.py` – Generates statistics and charts for all of France. 
- `regional.py` – Generates statistics for all regions of France.  
- `IDF.py` – Generates statistics for all departements of Île-de-France.  
- `Occitanie.py` – Generates statistics for all departements of Occitanie.  
- `DROM.py` – Generates statistics for French overseas regions (DROM).  

Each script reads Excel files containing the relevant data (not included in the repository) and updates a PowerPoint template by filling slides with the calculated statistics and charts.

## Dependencies

- Python 3.x  
- pandas  
- python-pptx
- openpyxl

You can install the required packages with:

```bash
pip install pandas python-pptx openpyxl
```

## Usage

```bash
python national.py
python regional.py
python IDF.py
python Occitanie.py
python DROM.py
```

The scripts will generate an updated PowerPoint with the latest statistics and charts.

## Notes

- The Excel files used as input are confidential and are not included in this repository.  
- The PowerPoint template is included to demonstrate the layout, but it does not contain real statistics.

## Author

[Yasmine Aissa/Yasmine56] 
