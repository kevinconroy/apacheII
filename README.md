## Acute Physiology and Chronic Health Evaluation II (APACHEII)

This is a VBA function for Microsoft Excel that calculates the APACHEII score for a patient. This score estimates ICU mortality. You should use the worst value for each physiological variable within the past 24 hours.

This macro was written using the formula provided by Knaus WA et al (1985). This function ensures that all values are within a sane range. Upper and lower limits allow for some "world record" setting conditions and might suggest conditions beyond known survivable ranges. These are meant to ensure that there are not obvious coding or data entry errors.

References:
* https://en.wikipedia.org/wiki/APACHE_II
* http://www.sfar.org/scores2/apache22.html - Used for quality assurance of macro below
* Knaus WA et al. APACHE II : A severity of disease classification system. Crit Care Med. 1985;13:818-2 https://www.ncbi.nlm.nih.gov/pubmed/3928249


## How to Use

### Example File
1. Download APACHE-II-Example.xlsm
1. Be sure to "Enable Macros"
1. Edit the data or add rows to calculate APACHE II scores.


### Adding Macro to Existing Excel Workbook

1. Open a new or existing Excel file.
2. "Save As..." and select the format "Excel Macro-Enabled Workbook (.xlsm)"
3. Go to Tools -> Macro -> Visual Basic Editor
4. Copy and paste the contents of the 'APACHE-II.bas' file into the editor that appears. You can alternative "Import" this file.
5. Save the workbook.
6. Close the Visual Basic Editor and go back to your workbook.
7. Use the APACHEII function to calculate the APACHEII score for the given arguments
8. Once you have an APACHE II score, you can use APACHEII_DEATHRATE or APACHEII_DEATHRATE_ADJUSTED to determine the estimated death rate. See http://www.sfar.org/scores2/apache22.html for adjustment factors.
9. Remember to "Enable Macros" each time you open the file.

