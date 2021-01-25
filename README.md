# PowerpointToPDF PresentationToPDF converts a Microsoft Powerpoint presentation to a PDF file containing one page\ for each animation step (i.e. each mouse-click).

To do this, it temporarily marks the original slides as "hidden", then creates one new slide for the end-state of each animation step. Both the old and new slides are tagged to identify which are which, and whether or not any of the original slides were already hidden. The result is "printed" to the "Microsoft Print to PDF" pseudo-printer, which is installed as part of Windows 10. The user then supplies a filename to a dialogue generated by the printer driver. Finally, the generated slides are deleted, the original slides unhidden and tags removed, restoring the presentation to its original state.

Optionally, slide numbers (in format *\<old slide number\>**-**\<animation step\>*) can be added to the result.

It is also possible to suppress the "restore" function and edit the generated slides. This is to allow the user to manually fix any parts of the conversion that the automatic process haven't worked properly (I'm not aware of any, but I'm sure there are some!)

When the macro is run on a presentation that has already been converted, but not restored, it will warn the user and restore the presentation to normal.

TECHNICAL NOTE: To achieve this, shapes that are not visible at a given step are deleted from the generated slide. Shapes containing animated paragraphs are left in place, but those paragraphs that are not visible within them at a given step are camouflaged rather than deleted, i.e. the font colour is changed to match the shape colour. This ensures other, visible paragraphs do not move or increase in size within the shape.
