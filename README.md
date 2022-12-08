# VBA_BOM_Compare_SAP_CAD

## Motivation
At a previous workplace, the accounting program (SAP) was being manually updated with input from the CAD development program (Inventor). 
This led to costly errors, ordering outdated parts and assemblies. 

## Methods
To prevent such errors in the future, I wrote a VBA program to export Bills of Materials from SAP and CAD to perform an automatic comparison. 
I chose VBA to seamlessly integrate into a platform already available at the company. 
Due to restricted SAP scripting access, an automated SAP_BOM export is achieved by simulated keystrokes. 

## Results
To identify and meet user requirements I went through several iterations with my colleagues. 
During these iteration steps, the user interface was extended by:
- Display separate lists to show IDs which are only present in either CAD or SAP
- Display differences in the assembly tree
- Option to ignore certain number circles, e.g. for raw material only shown in SAP or subcomponents only included in CAD
- Option to import several sets of BOM raw data for multiple comparisons

## Conclusion
I am happy to know that this program is still in use even 2+ years after I left the company on 30.09.2020 due to a Covid savings initiative. 
