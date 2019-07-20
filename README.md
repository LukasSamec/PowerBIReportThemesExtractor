# PowerBIReportThemesExtractor

This simple program made in C# allows you to extract [theme](https://docs.microsoft.com/en-us/power-bi/desktop-report-themes) from Power BI file (.pbix) or Power BI template file (.pbit).


## How to use

All you need to do is set absolute path to Power BI file (.pbix) or Power BI template file (.pbit) and absolute path to where the extracted theme will be saved. You can enter the paths manually or you can use open / save file dialogs that will help you to set the paths to the files.

## Limitations

* To work properly all colors in Power Bi must defined via HEX code (for example: #000000). To achieve this all colors in Power Bi   must be set up by "Custom color" function in Power BI in color picking. If the program finds color which is not defined by HEX code, the color will be set to default value which the user can define before extraction.

* Power Bi report theme can contain definitions only for types of visual, so in the theme there can not be definition for the same type of visual multiple times. If the program finds multiple definition for the same visual (for example multiple pie chart), the program will extract definition of the first occurence.
