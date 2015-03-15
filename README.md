h1. Acute Physiology and Chronic Health Evaluation II (APACHEII)

This is a VBA function for Microsoft Excel that calculates the APACHEII score for a patient. This score estimates ICU mortality.

Medical Professionals: You should use the worst value for each physiological variable within the past 24 hours.

This macro was written using the formula provided by Knaus WA et al (1985).
This function ensures that all values are within a sane range. Upper and lower limits allow for some "world record" setting conditions
and might suggest conditions beyond known survivable ranges. These are meant to ensure that there are not obvious coding or data entry errors.

References:
* [https://en.wikipedia.org/wiki/APACHE_II]
* [http://www.sfar.org/scores2/apache22.html] - Used for quality assurance of macro below
* Knaus WA et al. APACHE II : A severity of disease classification system. Crit Care Med. 1985;13:818-2 [https://www.ncbi.nlm.nih.gov/pubmed/3928249]


h1. Usage

# Open a new or existing Excel file.
# Save the file using the "Excel Macro-Enabled Workbook (.xlsm)" format
# Tools -> Macro -> Visual Basic Editor
# Copy and paste the contents of the 'apacheII.vba' file into the editor that appears
# Save the file
# Close the Visual Basic Editor and go back to your workbook
# Use the APACHEII function to calculate the APACHEII score for the given arguments
# Once you have an APACHE II score, you can use APACHEII_DEATHRATE or APACHEII_DEATHRATE_ADJUSTED to determine the estimated death rate.

