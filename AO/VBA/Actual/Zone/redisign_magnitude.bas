Sub redig()
Dim f_c&, f_i&, f_y&, f_x&
Dim f_val as Variant
Dim LastColum&, LastRow&, start_row&, start_column&
Dim ar_Data()
Dim nm_brand$, key_brand$
Dim dicBrand as Variant, dicData as Variant
Dim colBrandRow as Collection
Dim objBrand as mySplitClients
Dim sts_data_row as Boolean

Set colBrandRow = New Collection

Set dicBrand = CreateObject("Scripting.Dictionary")
With dicBrand
    .Add "Total doors PPD", "PPD"
    .Add "Kérastase", "KR" 
    .Add "Redken", "RD" 
    .Add "Matrix", "MX" 
    .Add "Shu Uemura Prof.", "SU" 
    .Add "Essie Prof.", "ES" 
    .Add "Decleor", "DE" 
    .Add "Carita", "CR"
    .Add "Kéraskin", "KS"
End With

For Each sh in ThisWorkbook.Worksheets
    LastColum   = getLastColumn
    LastRow     = getLastRow
    
    sts_data_row = False
    key_brand = Empty
    Set objBrand = New mySplitClients
    For f_cl = 1 to LastColum
        For f_rw = 1 to LastRow
            f_val = Trim(Cells(f_rw, f_cl))
            If  dicBrand.Exists(f_val) and key_brand <> f_val Then
                key_brand = f_val
                objBrand.
                colBrandRow.add objBrand, Key:=key_brand
                Set objBrand = New mySplitClients
                name = dicBrand.Key(f_val)
                intCountSalons = Empty
                intCountHaircareSalons = Empty
                intCountSkincareSalons = Empty
                intCountNailSalons = Empty
                intCountColoxSalons = Empty
            End If

            If dicBrand.Exists(key_brand) Then
                sts_data_row = True
                Select Case f_val
                    Case "PPD doors - direct", "Buying salons - direct"
                        objBrand.intCountSalons = f_rw
                    Case "of which Haircare"
                        objBrand.intCountHaircareSalons = f_rw
                    Case "of which Skincare"
                        objBrand.intCountSkincareSalons = f_rw
                    Case "of which Nail"
                        objBrand.intCountNailSalons = f_rw
                    Case "of which Salons Colox - direct"
                        objBrand.intCountColoxSalons = f_rw
                End Select
                colBrandRow 
            End If
            Next f_rw
        Next f_cl
                   
Next sh


End Sub


