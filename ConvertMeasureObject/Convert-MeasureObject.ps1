function Convert-MeasureObject {
    <#
    .SYNOPSIS
    Convert the output of Measure-Oject into a flattened object with each measured property being it's own property.

    .DESCRIPTION
    Convert the output of Measure-Oject into a flattened object, each measured property becomes a property of the
    new object generated.

    This will convert output from Measure-Object that initially looks like this:

        Average       Sum Maximum Minimum Property
        -------       --- ------- ------- --------
                    123.45                 Interest
                123456.78                 Total

    The resulting object will look like this:

            Total Interest
            ----- --------
        123456.78   123.45

    .PARAMETER InputObject
    Object from Measure-Object. Accepts pipeline input or a general array of objects with a single shared property.

    .PARAMETER Property
    Property from Measure-Object that will be used for the conversion and flattening of the object.

    .EXAMPLE
    $SomeObject | Measure-Object -Sum | Convert-MeasureObject -Property 'Sum'

    This will convert the output from Measure-Object based on the content of the Sum property.
    #>
    [cmdletbinding()]
    param (
        [Parameter(ValueFromPipeline)]
        [Object[]]$InputObject,

        [string]$Property
    )

    begin {
        $Hash = @{}
    }
    process {
        foreach ($row in $InputObject) {
            $Hash.add($Row.Property, $InputObject.Where({$_.Property -eq $row.property}).$Property)
        }
    }
    end {
        [PsCustomObject]$Hash
    }
}
