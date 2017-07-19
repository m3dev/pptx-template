# Specification of model

pptx-template reads "model" file which contains data to be imported into the presentation.

A model file can be written in both of JSON and xlsx(Excel) format. When the model file extension is 'xlsx', pptx-template reads model file as 'Excel' format, otherwise reads as 'JSON' format.

## Example of JSON
```
{
  "slides:" {    
                 // All slides settings should be put under "slides" element.
                 // Each slides should have its own "{id:foobar}" text frame somewhere in the slide.

                 // for slide "{id:p0}"
      "p0": "remove",  //  "remove" means to remove the slide

                 // for slide "{id:p1}"
      "p1": {
        "greeting": "Hello",          // simple substituion for "{greeting}"
        "num": [ "100", "200" ],      // values can be structured as an array or a dictionary(hash)
                                      // ex. "{num.0}" refers "100"
        "t": [                        // Nested array is useful for values in a table.
          ["Year", "Sales", "Cost"],  // | {t.0.0} | {t.0.1} | {t.1.1} |
          [2016, 100, 200 ],          // | {t.1.0} | {t.1.1} | {t.1.2} |
          [2017, 150, 180 ]           // | {t.2.0} | {t.2.1} | {t.2.2} |
        ]
      },

                 // for slide "{id.p2}"
      "p2": {
        "greeting": "Hola",
        "chart0": {                           // A chart should include its key in its title, like "{chart0}MyChart"
                                              // A chart setting need its special structure as follows:
          "file_name": "data-for-chart-csv"   //  file_name : file name to be imported. Files should be ened with ".csv" or ".tsv"
          "body": "Year,Sales,Cost\n2001,200,150",
                                              //  body: The CSV contents itself instead of file name. Prior than "file_name"
          "tsv_body":  "Year\tSales\n2001\t200\t150",
                                              //  tsv_body: TSV contents. The priority is "body" > "tsv_body" > "file_name"
          "value_axis_max": 100,              //  value_axis_max: The maximum value for Y axis
          "value_axis_min": 200               //  value_axis_min: The minimum value for Y axis
        }
      }
  ]
}

```

## Example of Excel

Excel file should have a sheet named "model"

The "model" sheet should have one header line and value lines. The header line will be ignored.

Each line should have these items:

  - SlideId : ID for slide which contains ``{id:<slide_id>}`` text frame somewhere in the slide
  - Key     : Key for substitution, written as ``{<key>}`` in slides
  - Value   : Value for Substitution. Part of Excel's cell format will be applied.
  - Range   : Range for chart/table data. If Value exists, Range will be ignored.
  - Options : Export options to create range value. 

```
ex)
SlideId    Key          Value    Range            Options
----------------------------------------------------------
p1         greeting.en  Hello    (empty)
p1         greeting.es  Hola     (empty)
p1         chart        (empty)  chart1-range
p1         chart2       (empty)  =sheet2!A3:C10   Transpose,Array
```

#### Supported Excel formats for Value

  - Digits of fraction : ex) ``0.000`` 
    - only ``.`` can be used. (not ``,``)
  - Percent : value will be multiplied by 100 and followd by ``%``. ex) ``0.0%`` 
  - Not support date formats like ``yyyy-mm-dd``
  - Not support ``,`` for 3-digit delimiter
  
#### Range

There're two ways to specify data location:

  - Name of Excel's named range : ex) ``range1``
  - Range reference : ex) ``=A1:C99``
    - This will be shown '#VALUE!' in excel. Excel's ``FORMULATEXT()`` might help to check if the value is correct
    - More than 2 ranges can be written, followed by ``,``: ex) ``=A1:A10,C1:C10``
    - Only relative form is supported. Not supported ``=A$1:A$99`` nor ``=<named_range>``
    
#### Options

Options specifies how to export ranged data. Multiple values cen be written with separated by ``,``

  - SideBySide : To apply multiple ranges as they placed side-by-side, not top-to-bottom(default). 
  - Transpose  : To invert column and row
  - Array      : To Make array object to refer as ``{x.0.1}`` style key especially from table in slides.

