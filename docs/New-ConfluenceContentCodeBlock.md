---
external help file: tcs.confluence-help.xml
Module Name: tcs.confluence
online version:
schema: 2.0.0
---

# New-ConfluenceContentCodeBlock

## SYNOPSIS
Creates a Confluence code block macro.

## SYNTAX

```
New-ConfluenceContentCodeBlock [-Content] <String> [[-Language] <String>] [[-Theme] <String>] [-LineNumbers]
 [[-Collapse] <Boolean>] [<CommonParameters>]
```

## DESCRIPTION
New-ConfluenceContentCodeBlock returns the storage-format markup for the Confluence "code"
macro.
The code is placed in a CDATA section, so it is shown exactly as given; a "\]\]\>"
sequence in the code is split so it cannot end the CDATA section early.

## EXAMPLES

### EXAMPLE 1
```
New-ConfluenceContentCodeBlock -Content 'Get-Process' -Language powershell -LineNumbers
```

Returns a PowerShell code block macro with line numbers.

## PARAMETERS

### -Content
The code to show.

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: True
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Language
The language used for syntax highlighting, for example powershell, bash or json.
Default none.

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: 2
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Theme
The macro theme: Default, Midnight, Eclipse or Emacs.

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: 3
Default value: Default
Accept pipeline input: False
Accept wildcard characters: False
```

### -LineNumbers
Show line numbers.

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -Collapse
Collapse the code block when the page loads.

```yaml
Type: Boolean
Parameter Sets: (All)
Aliases:

Required: False
Position: 4
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

## OUTPUTS

### System.String
## NOTES

## RELATED LINKS
