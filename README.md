# Drawings Macro: MakeDrawings.swp
### SOLIDWORKS macro to create a drawing from an open assembly model
* MakeDrawings.swp is the complete macro, including the class and module listed below.
   - It was written using SOLIDWORKS 2026, but should also be backwards compatible.

### Helpful resources
* Utilities.bas is a module for changing filepath extensions. You must add Microsoft Scripting Runtime as a reference to your project to use this module.
* DrawingBuilder.cls is the class that does all of the work: 
  - create a new drawing
  - populate pre-defined views
  - add a BOM table from a template
  - Autoballoon a selected view
  - save the drawing
  - save as PDF
  - save the BOM as Excel
