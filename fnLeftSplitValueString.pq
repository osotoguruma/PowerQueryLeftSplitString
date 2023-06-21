let fnLeftSplitValueString = (pTable as text, pSourceString as text, pTargetValue as text, pTargetString as text)  =>
    
    let 

        Source = Excel.CurrentWorkbook(){[Name=pTable]}[Content],
        ForceTypeText = Table.TransformColumnTypes(Source,{{pSourceString, type text}}),

        Expr1="each Text.SplitAny(Text.Lower([" & pSourceString & "])," & """abcdefghijklmnopqrstuvwxyz""" & "){0}", 
        ExtractValue = Table.AddColumn(ForceTypeText, pTargetValue, Expression.Evaluate(Expr1, #shared)),

        Expr2="each Text.End(Text.From([" & pSourceString & "]), Text.Length([" & pSourceString & "])-Text.Length([" & pTargetValue & "]))",
        ExtractText = Table.AddColumn(ExtractValue, pTargetString, Expression.Evaluate(Expr2, #shared)),

        TrimBlank = Table.ReplaceValue(ExtractText," ","",Replacer.ReplaceText,{pTargetValue}),
        ForceTypeValue = Table.TransformColumnTypes(TrimBlank,{{pTargetValue, type number}})

    in
        ForceTypeValue
in 
    fnLeftSplitValueString