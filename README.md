Item Label =
VAR it    = DimItem[Item]
VAR p1    = SEARCH ( "(", it, 1, 0 )
VAR p2    = SEARCH ( ")", it, 1, 0 )
VAR inner = IF ( p1 > 0 && p2 > p1, MID ( it, p1 + 1, p2 - p1 - 1 ) )
VAR nmAll = LOOKUPVALUE ( EntityList[Entity], EntityList[Code], it )
VAR nmIn  = LOOKUPVALUE ( EntityList[Entity], EntityList[Code], inner )
RETURN
    IF ( NOT ISBLANK ( nmAll ), nmAll & " (" & it & ")",
        IF ( NOT ISBLANK ( nmIn ),
             SUBSTITUTE ( it, "(" & inner & ")", nmIn & " (" & inner & ")" ),
             it ) )


_________________________

Cash EoP calc =
CALCULATE (
    [Value Current],
    FactData[Topic] = "Remittance Holding",
    NOT ( LEFT ( FactData[Item], 3 ) = "o/w" || CONTAINSSTRING ( FactData[Item], "(EoP)" ) )
)


___________
