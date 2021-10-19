# StateImpactReportTool

Generates State Impact reports for generated CVRS Add, Update and Delete files.

---

This tool must be run separately for generating state impact report for Add/Update/Delete files. Meaning, if you want to generate a state impact report for Update and Delete files, you will have to run the tool twice. At a time, this tool can process more than one file, meaning, if the Update records are split across multiple files, put them in a folder and point the folder in the config mentioned below and the tool would generate one state impact report for the multiple update files.

• Unzip the StateReportTool.
• Open the StateReportTool.exe.config in a notepad and edit the following:

• Edit the value field (the highlighted field below) for the filePath in the config to point to the folder where the, e.g., Delete files are.
     o <add key="filePath" value="C:\TestFiles\DeleteFiles"/>
• Edit the value field (the highlighted field below) for the excelPath in the config. This is the folder and the file name (name of your choice ) for the State Impact Report.
     o <add key ="excelPath" value ="C:\TestFiles\GeneratedExcel\StateImpact_UD.xls"/>

• Double click the StateReportTool.exe.
• Once the processing is done, the state impact report should be available in the path mentioned in the excelPath in the StateReportTool.exe.config
