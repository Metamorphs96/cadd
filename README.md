# cadd
Experiments with CADD Systems. In particular experiments with AutoCAD LT command scripts
Alone examples of using IntelliCAD Lisp and COM automation
IntelliCAD Lisp typically requires little change to run in AutoLISP
The ActiveX/COM/.Net Object models are not the same and significant amount of modification is needed to convert to run in AutoCAD.

# Automation, Drawing Creation
Note there is typically no need for AutoLISP. Unlike other CAD programs, AutoCAD has a command language and this language can be scripted. Whilst the scripts themselves do not have enhancements, and cannot do calculations or have flow control, the scripts can be generated by other programs. The script generator can be written in any suitable programming language, if that language is also suitable for COM/.NET then if the program is written in a suitable modular and independent manner, it can be relatively easy to swap one module in and another out, so as to control a different CAD package. For example the script writing routines modified from AutoCAD to ProgeCAD, or to LibreCAD, or change from scripting to COM/.NET automation in AutoCAD, ProgeCAD, or DesignCAD (current versions appear to have reverted to BasicCAD).

Note that here automation is not considered to include enhancements or augmentation of actvity in the drawing editor. 

# Data Extraction
Data extraction of the primary tables sysvars, linetypes, text styles, dimension styles, view, UCS, viewports, in typically possible from the command line. Block attributes can be extracted using ATTEXT and a template. Entity data can be extracted by exporting a DXF file, and given the availability of open source, chances are there is a suitable library available. Though for many requirements only a simple parser is required not a full parser, just write to extract that which is needed and ignore the rest.

# Interaction
If objective is interaction, then need to customise menus and toolbars/ribbons with command macros, which can pause, and otherwise the use of DIESEL macros. If want to do more then would need AutoLISP or ActiveX/COM/.NET or other application program interface (API).

#Conclusion
A lot of AutoLISP programs were written because AutoLISP was available, and the author could use, not because such language was required to automatically accomplish a task in AutoCAD. So if you are an AutoCAD LT user then need to think more carefully about requirements than leap to conclusion not possible because don't have AutoLISP.

As for inclusion of AutoLISP in AutoCAD LT 2024, no real benefit if needs are drawing creation and data extraction. If requirement is interaction with the drawing editor, then it is probably more important to ask the question: is AutoCAD the appropriate technology in the first place?

