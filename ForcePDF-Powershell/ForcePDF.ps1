# creates a "drive" to access the HKCR (HKEY_CLASSES_ROOT)
New-PSDrive -Name HKCR -PSProvider Registry -Root 
HKEY_CLASSES_ROOT

If ('HKCR:\.pdf')
{
    # this is the .pdf file association string
    $PDF = 'HKCR:\.pdf'
    New-ItemProperty $PDF -Name NoOpenWith
    New-ItemProperty $PDF -Name NoStaticDefaultVerb
}

If ('HKCR:\.pdf\OpenWithProgids')
{
    # this is the .pdf file association string
    $Progids = 'HKCR:\.pdf\OpenWithProgids'
    New-ItemProperty $Progids -Name NoOpenWith
    New-ItemProperty $Progids -Name NoStaticDefaultVerb
}
