public Sub CalculoCargos(byval empleo as long, paramarray lineasCargo() As Variant)
    dim DatoscargosORI As String, RstCargosORI as recordset
    DatoscargosORI = ExpandirSQL("SELECT CreditoID, SucursalCod, CreditoCuenta, CreditoFechaProxVto, CreditoSdoCap, EstadoCreditoID, OficinaCod, empleoid " & _
    "FROM TempConsCredAlta " & _
    "WHERE EstadoCreditoID IN (4,8,10) " & _
    "AND OficinaCod IN {1}" & _
    "AND empleoid = {0}",  empleo, lineasCargo())

    Set RstCargosORI = OpenRS(DatoscargosORI)
    RunSQL ("DELETE * FROM TempGestionCobranza")

    If Not RstCargosORI.EOF Then
        Do Until RstCargosORI.EOF
            LiquidaBoletaCob RstCargosORI!CreditoID, 4, 100, 0, Now
            GraboComprobanteCargo True, RstCargosORI!CreditoID
            RstCargosORI.MoveNext
        Loop
    End If
    RstCargosORI.Close
    MuestroDatosCargo(empleo, lineasCargo())
end sub

public sub MuestroDatosCargo(byval empleo as long, paramarray lineasCargo() as Variant)
    dim datoscargos As String, rstcargos as recordset
    datoscargos = expandirSQL("SELECT a.CreditoID, a.SucursalCod, a.CreditoCuenta, a.CreditoFechaProxVto, a.CreditoSdoCap, a.EstadoCreditoID, a.OficinaCod, a.empleoid, Sum(b.ImporteConcepto) AS SumaDeImporteConcepto, b.ComprobanteID " & _
                "FROM TempConsCredAlta AS a INNER JOIN TempGestionCobranza as b ON a.CreditoID = b.CreditoId " & _
                "GROUP BY a.CreditoID, a.SucursalCod, a.CreditoCuenta, a.CreditoFechaProxVto, a.CreditoSdoCap, a.EstadoCreditoID, a.OficinaCod, a.empleoid, b.ComprobanteID " & _
                "HAVING OficinaCod IN {1} AND a.empleoid = {0}", empleo, lineasCargo())

    Set RstCargos = OpenRS(datoscargos)
    
    '******************* Armo cargos completos *****************
    PasoATraves datoscargos, "TempoCargosCred"
    
    
    If not RstCargos.eof Then
        Forms![002AltaCreditos]![MuestroCargo].Form.RecordSource = datoscargos
        Forms![002AltaCreditos]!Cargo = DSum("SumaDeImporteConcepto", "TempoCargosCred")
    Else
        Forms![002AltaCreditos]![MuestroCargo].Form.RecordSource = ""
        Forms![002AltaCreditos]!Cargo = 0
    End If
    
    RstCargos.Close
end sub
