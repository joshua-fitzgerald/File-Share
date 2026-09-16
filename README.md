Run it like this:

powershell
.\Prepare-ITBoostForITGlue.ps1 `
  -ZipPath "C:\Downloads\6aa32c1eec08039e4bd008f6_1789078560596.zip" `
  -OutputFolder "C:\Downloads\itglue-output"

It will then show you something like:

Found 3 categories in the export:

  [1] global  (47 documents)
  [2] Client Knowledgebase  (12 documents)
  [3] _HNTFB Public facing T&C's  (5 documents)

How would you like to output the ZIPs?
  [M] Merge all categories into one ZIP
  [S] Separate ZIP per category

Enter M or S:

If you choose S, you'll get one ZIP per category named itglue-import-global.zip, itglue-import-Client Knowledgebase.zip, etc. — handy if you want to import each category into a different IT Glue org or folder.
