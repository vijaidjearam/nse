# This method shows automation via Xlwings UDFs

Please install the following requirements via pip

```batch
pip install xlwings
```

One-time Excel preparations

Enable *Trust access to the VBA project object model* under *File > Options > Trust Center > Trust Center Settings > Macro Settings*. You only need to do this once. Also, this is only required for importing the functions, i.e. end users won’t need to bother about this.

Install the add-in via command prompt: 

```batch
xlwings addin install
```

Workbook preparation
The easiest way to start a new project is to run on a command prompt the following command:

```batchs
xlwings quickstart myproject
````

This automatically adds the xlwings reference to the generated workbook.

Note: The excel should be saved in .xlsm (Macros)

If you have created a project already created using xlwings udfs you can use it right now, as long as the excel file and python file are in the same location.
