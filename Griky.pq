let
    Origen = Excel.CurrentWorkbook(){[Name="Tabla5"]}[Content],
    #"Columnas quitadas" = Table.RemoveColumns(Origen,{"fecha_de_corte", "Nivel", "Transaccional", "Saldo inicial", "Movimiento débito", "Movimiento crédito", "Saldo final"}),
    #"Columnas reordenadas" = Table.ReorderColumns(#"Columnas quitadas",{"empresa", "fiscal_month", "fiscal_year", "Código cuenta contable", "Nombre cuenta contable", "Identificación", "Nombre tercero", "statement", "class", "group", "gross_to_net", "fin_account_id", "fin_account", "fin_subaccount", "dato_financiero", "dato_financiero_USD"}),
    #"Valor reemplazado" = Table.ReplaceValue(#"Columnas reordenadas","ALIANZA SUPERIOR SAS","ALIANZA",Replacer.ReplaceText,{"empresa"}),
    #"Valor reemplazado1" = Table.ReplaceValue(#"Valor reemplazado","EDTECH4LIFE SL","EDTECH4LIFE",Replacer.ReplaceText,{"empresa"}),
    #"Valor reemplazado2" = Table.ReplaceValue(#"Valor reemplazado1","HOLOVERSITY SAS","HOLOVERSITY",Replacer.ReplaceText,{"empresa"}),
    #"Valor reemplazado3" = Table.ReplaceValue(#"Valor reemplazado2","VIRTUOSITY 11.11 LLC","VIRTUOSITY",Replacer.ReplaceText,{"empresa"}),
    #"Tipo cambiado1" = Table.TransformColumnTypes(#"Valor reemplazado3",{{"Código cuenta contable", type text}}),
    #"Personalizada agregada" = Table.AddColumn(#"Tipo cambiado1", "Largo", each Text.Length([Código cuenta contable])),
    #"Columnas reordenadas1" = Table.ReorderColumns(#"Personalizada agregada",{"Largo", "empresa", "fiscal_month", "fiscal_year", "Código cuenta contable", "Nombre cuenta contable", "Identificación", "Nombre tercero", "statement", "class", "group", "gross_to_net", "fin_account_id", "fin_account", "fin_subaccount", "dato_financiero", "dato_financiero_USD"}),
    #"Valor reemplazado4" = Table.ReplaceValue(#"Columnas reordenadas1","Balance Sheet ","Balance General",Replacer.ReplaceText,{"statement"}),
    #"Valor reemplazado5" = Table.ReplaceValue(#"Valor reemplazado4","Balance Sheet","Balance General",Replacer.ReplaceText,{"statement"}),
    #"Columnas reordenadas2" = Table.ReorderColumns(#"Valor reemplazado5",{"Largo", "statement", "empresa", "fiscal_month", "fiscal_year", "Código cuenta contable", "Nombre cuenta contable", "Identificación", "Nombre tercero", "class", "group", "gross_to_net", "fin_account_id", "fin_account", "fin_subaccount", "dato_financiero", "dato_financiero_USD"}),
    #"Personalizada agregada1" = Table.AddColumn(#"Columnas reordenadas2", "region", each if [empresa] = "EDTECH4LIFE" then "EUROPA"
else "LATAM"),
    #"Columnas reordenadas3" = Table.ReorderColumns(#"Personalizada agregada1",{"Largo", "statement", "empresa", "region", "fiscal_month", "fiscal_year", "class", "fin_account_id", "group", "gross_to_net", "fin_account", "fin_subaccount", "Código cuenta contable", "Nombre cuenta contable", "Identificación", "Nombre tercero", "dato_financiero", "dato_financiero_USD"}),
    #"Tipo cambiado2" = Table.TransformColumnTypes(#"Columnas reordenadas3",{{"dato_financiero_USD", Currency.Type}, {"dato_financiero", Currency.Type}}),
    #"Poner En Mayúsculas Cada Palabra" = Table.TransformColumns(#"Tipo cambiado2",{{"Nombre tercero", Text.Proper, type text}}),
    #"Tipo cambiado3" = Table.TransformColumnTypes(#"Poner En Mayúsculas Cada Palabra",{{"Identificación", type text}}),
    #"Columnas reordenadas4" = Table.ReorderColumns(#"Tipo cambiado3",{"Largo", "statement", "empresa", "region", "fiscal_month", "fiscal_year", "class", "fin_account_id", "group", "gross_to_net", "fin_account", "fin_subaccount", "Nombre tercero", "Identificación", "Código cuenta contable", "Nombre cuenta contable", "dato_financiero", "dato_financiero_USD"})
in
    #"Columnas reordenadas4"