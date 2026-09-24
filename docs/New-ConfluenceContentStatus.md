---
external help file: tcs.confluence-help.xml
Module Name: tcs.confluence
online version:
schema: 2.0.0
---

# New-ConfluenceContentStatus

## SYNOPSIS
Creates a Confluence status lozenge (status macro).

## SYNTAX

```
New-ConfluenceContentStatus [-Text] <String> [[-Colour] <String>] [-Subtle]
 [-ProgressAction <ActionPreference>] [<CommonParameters>]
```

## DESCRIPTION
New-ConfluenceContentStatus returns the storage-format markup for the Confluence "status"
macro: a small coloured label such as DONE or IN PROGRESS.
The text is escaped.
Use the result
in a paragraph or, with New-ConfluenceContentTable -Raw, in a table cell.

## EXAMPLES

### EXAMPLE 1
```
New-ConfluenceContentStatus -Text 'Done' -Colour Green
```

Returns \<ac:structured-macro ac:name="status"\>\<ac:parameter ac:name="colour"\>Green\</ac:parameter\>\<ac:parameter ac:name="title"\>Done\</ac:parameter\>\</ac:structured-macro\>.

## PARAMETERS

### -Text
The text of the lozenge.
Confluence shows it in capitals.

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

### -Colour
The colour: Grey (default), Red, Yellow, Green, Blue or Purple.
-Color is an alias.

```yaml
Type: String
Parameter Sets: (All)
Aliases: Color

Required: False
Position: 2
Default value: Grey
Accept pipeline input: False
Accept wildcard characters: False
```

### -Subtle
Use the subtle (outlined) style instead of a filled lozenge.

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

### -ProgressAction
{{ Fill ProgressAction Description }}

```yaml
Type: ActionPreference
Parameter Sets: (All)
Aliases: proga

Required: False
Position: Named
Default value: None
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
