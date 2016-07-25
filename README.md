# TactileCheck
Tactile check template #word 2013 #VBA #docm

This word template is used during my masterthesis research.
to start with, I am not a profesional software developper just an enthosiastic trail and error get thing working code creator. 
The script is intended to annotate weak and strong phrases in the text 'body'
The phrases are loaded from 2 tables in the template
Expansion of the tables to accomodate more phrases (row) or categories (column) is possible

The annotation is based on the standard word Search functionality and creates for each found phrase an standard word Note.
The note includes an ID, category and the phrase.
Each phrase found is also highlighted in the text, font size is increased (+)
and the typeface is changed from bold for stron and itallic for weak phrases.

Code is included to export all Notes to Excel for further analysis

The code is not writen for efficiency, for its purpose it is fast enough.

Set-up:
copy load_table_from_doc_v0010-template - Weak - Strong workflow.dotm to your Custom Office Templates location.
Create a new tactile check document based on the dotm template.
(or open the included example)

additional Reference settings if runtime errors are pressented (depending on local installation)
Enable "DEVELOPER"  => and open Visual Basic Editor (ALT + F11)
Select 'Tools' => References
ensure that the following settings are enabled:
-Visual Basic For Applications
-OLE Automation
-Microsoft Word 15.0 Object Library
-Microsoft Office 15.0 Object Library
-Microsoft Excel 15.0 Object Library
-Microsoft Forms 2.0 Object Library


