
Public Sub TheOneTrueRutine(LineaId As Integer, Capital As Currency, Plazo As Integer, difDias1eraCuota)
    

On Error GoTo ErrFrances

Dim RstLinea As Recordset
Dim Tempo As Recordset
Dim RstCargosORI As Recordset
Dim RstCargos As Recordset
Dim Impuesto As Double
Dim TasaX As Double
Dim Vartasa As Single
Dim VarConceptoCuota As Single
Dim VarFondoGarantia  As Single
Dim VarSeguro As Single
Dim IVASobreInteres As Single
Dim CapitalPrestado As Currency, CapitalSolicitado As Currency, CapitalPrestadoAjustado As Currency
Dim Seguro As Currency, AjustePrimerCuota As Currency
Dim SaldoCapital As Currency
Dim Interes As Currency
Dim VarSeguroVida As Currency
Dim VarFondoGtia As Currency
Dim datoscargos As String
Dim DatosRetenciones As String
Dim CuotaConIva As Currency, CuotaSinIva As Currency
Dim Retenciones As Currency
Dim FechaCuota As Date
Dim i As Integer
Dim DiasAjuste As Integer


Dim SelladoII, SeguroII 'REVISAR SI SE USAN


Set RstLinea = OpenRS("SELECT * FROM LineasCreditos WHERE LineaCreditoID = {0}", LineaId)
If Not RstLinea.EOF Then
    
    Impuesto = 1 + (RstLinea!LineaIVAGeneral / 100)
    IVASobreInteres = 1 + (RstLinea!LineaIVASobreInt / 100)
    Vartasa = DameTasa(LineaId, Plazo)
    CapitalSolicitado = Capital
    DiasAjuste = difDias1eraCuota
    
    'Calculo Ajuste 1º cuota
    If difDias1eraCuota > 0 Then
          AjustePrimerCuota = ((Vartasa / 360) * difDias1eraCuota * CapitalSolicitado) / 100
    End If
    
    VarSeguro = RstLinea!LineaSeguroVida
    VarFondoGarantia = RstLinea!LineaFdoGarantia
    
    If Forms![002AltaCreditos]!AniosFin < 70 Then
        VarFondoGarantia = 0
    Else
        VarSeguro = 0
    End If
     
    If RstLinea!LineaSellado > 0 Then
        SelladoII = RstLinea!LineaSellado / Plazo
    End If
    
    If RstLinea!LineaSeguroVida > 0 Then
        SeguroII = RstLinea!LineaSeguroVida / Plazo
    End If
    
    DoCmd.RunSQL ("DELETE Liquidacion.* FROM TempoLiquidacion")
    Set Tempo = openrs("TempoLiquidacion") 
    
    
    CapitalPrestado = Capital
    SaldoCapital = CapitalPrestado
    TasaX = Vartasa / 12 / 100
    Capital = 0
    Interes = 0
    FechaCuota = Det1erVTO(LineaId, Date)
    CuotaConIva = Pmt(TasaX * IVASobreInteres, Plazo, -CapitalPrestado, 0, 0)
    CuotaSinIva = Pmt(TasaX, Plazo, -CapitalPrestado, 0, 0)
    SaldoCapital = CapitalPrestado
    
    
    
    '? CALCULO DEL CUADRO DE MARCHA DEL CREDITO
    '  Ajuste 1º cuota
    If difDias1eraCuota > 0 Then
        Tempo.AddNew
        Tempo!Ncuota = 0
        Tempo!fechavto = Format(DateAdd("d", difDias1eraCuota, Now), "dd/mm/yyyy")
        Tempo!Tasa = Vartasa
        Tempo!TasaNominalAnual = Vartasa
        Tempo!Solicitado = CapitalPrestado
        Tempo!Prestado = CapitalPrestado
        Tempo!Capital = 0
        Tempo!Interes = AjustePrimerCuota
        Tempo!IvaInteres = AjustePrimerCuota * RstLinea!LineaIVASobreInt / 100
        Tempo!SdoCap = CapitalPrestado
        Tempo!SeguroVida = CapitalPrestado * 0.001
        Tempo!FondoGarantia = 0
        Tempo!Gastos = 0 '(Capital + Interes) * (RstLinea!LineaGastosAdmin / 100)
        Tempo!IVAGastos = Tempo!Gastos * IVAGeneral / 100
        Tempo!Cuota = 0 'AjustePrimerCuota + Tempo!SeguroVida ' Redondeo2(Capital + Tempo!Interes + Tempo!IvaInteres + Tempo!Gastos + Tempo!IVAGastos + Tempo!SeguroVida)
        Tempo!linea = Forms![002AltaCreditos]!ListaLineasHabilitadas.Column(1)
        If Not IsMissing(Forms![002AltaCreditos]![002subaltacreditossolicitante]!NombreII) Then
            Tempo!Cliente = Forms![002AltaCreditos]![002subaltacreditossolicitante]!NombreII
            Tempo!Documento = Forms![002AltaCreditos]![002subaltacreditossolicitante]!NumeroDoc
        End If
        Tempo!Sellado = 0 
        Tempo.Update
    End If
        
    For i = 1 To Plazo
        Interes = SaldoCapital * TasaX
        InteresConIVA = Interes * IVASobreInteres
        Capital = CuotaSinIva - Interes
    
        If i = Plazo Then
            Capital = SaldoCapital
        End If    

        SaldoCapital = (SaldoCapital - Capital)
        VarSeguroVida = SaldoCapital * RstLinea!LineaSeguroVida / 1000
        VarFondoGtia = (SaldoCapital / 1000) * VarFondoGarantia
        'Grabo en la tempo
        Tempo.AddNew
        Tempo!Ncuota = i
        Tempo!fechavto = DameProximoDiaHabil(FechaCuota)
        Tempo!Tasa = Vartasa
        Tempo!TasaNominalAnual = Vartasa
        Tempo!Solicitado = CapitalPrestado
        Tempo!Prestado = CapitalPrestado
        Tempo!Capital = Capital
        Tempo!Interes = Interes
        Tempo!IvaInteres = Interes * RstLinea!LineaIVASobreInt / 100
        Tempo!SdoCap = SaldoCapital
        Tempo!SeguroVida = VarSeguroVida
        Tempo!FondoGarantia = VarFondoGtia
        Tempo!Gastos = (Capital + Interes) * (RstLinea!LineaGastosAdmin / 100)
        Tempo!IVAGastos = Tempo!Gastos * RstLinea!LineaIVAGeneral / 100
        Tempo!Cuota = Capital + Tempo!Interes + Tempo!IvaInteres + Tempo!Gastos + Tempo!IVAGastos
        Tempo!linea = Forms![002AltaCreditos]!ListaLineasHabilitadas.Column(1)

        If Not IsMissing(Forms![002AltaCreditos]![002subaltacreditossolicitante]!NombreII) Then '! poner un punto de debug y comprobar si siempre se entra aqui
            Tempo!Cliente = Forms![002AltaCreditos]![002subaltacreditossolicitante]!NombreII
            Tempo!Documento = Forms![002AltaCreditos]![002subaltacreditossolicitante]!NumeroDoc
        End If
        Tempo!Sellado = SelladoII
        Tempo.Update
        FechaCuota = Det1erVTO(LineaId, Date, i)
    Next

    '? APLICACION DE RETENCIONES
    
    RunSQL "DELETE * FROM RetencionesTempo"
  
    AgregarRetencion(2, CapitalPrestado, 0)
  
    If RstLinea!LineaFondocredito > 0 Then
    AgregarRetencion(22, CapitalPrestado * 0.015, Retenciones) ' Gastos Originacion
    AgregarRetencion(47, DSum("[SeguroVida]", "TempoLiquidacion"), Retenciones) ' Fondo Seguro
    End If
    
    If RstLinea!LineaSellado > 0 Then
    AgregarRetencion(24, (DSum("[Capital]", "TempoLiquidacion") + DSum("[Interes]", "TempoLiquidacion")) * 0.01, Retenciones) 'Sellado
    End If
    
    Forms![002AltaCreditos].Retenciones = Retenciones
    Forms![002AltaCreditos].CapitalIII = CapitalPrestado 
    
        
    If Forms![002AltaCreditos].DiasHastaInicio < 0 Then
        Forms![002AltaCreditos].DiasHastaInicio = 0
    End If
        
    Forms![002AltaCreditos].ImporteCuota = DLookup("Cuota", "TempoLiquidacion", "Ncuota = 1")
    Forms![002AltaCreditos].MuestroRetenciones.SourceObject = "SubMuestroRetenciones"
        
    DatosRetenciones = "SELECT RetencionesTempo.CreditoId, RetencionesTempo.CodigoConcepto, RetencionesTempo.ConceptoCuotaID, RetencionesTempo.CreditoId, RetencionesTempo.TempoImporte, RetencionesTempo.Dias, RetencionesTempo.PcIP, ConceptosCuota.ConceptoCuotaDescrip " & _
                        "FROM ConceptosCuota INNER JOIN RetencionesTempo ON ConceptosCuota.ConceptoCuotaID = RetencionesTempo.ConceptoCuotaID " & _
                        "WHERE (((RetencionesTempo.CodigoConcepto)=1) AND ((RetencionesTempo.ConceptoCuotaID)>2) AND ((RetencionesTempo.CreditoId)=0) AND ((RetencionesTempo.PcIP)='" & Trim(Forms!Menu!IPPc) & "'))"
    
    Forms![002AltaCreditos]![MuestroRetenciones].Form.RecordSource = DatosRetenciones
        

    VarConceptoCuota = IIf(Forms![002AltaCreditos].AniosFin < 70, 4,5) '! ENTIENDO QUE ESTA VARIABLE NO SE USA PARA NADA AQUI


    '? APLICACION DE CARGOS ABAJO
    ' DatoscargosORI = "SELECT CreditoID, SucursalCod, CreditoCuenta, CreditoFechaProxVto, CreditoSdoCap, EstadoCreditoID, OficinaCod, empleoid " & _
    '                     "FROM TempConsCredAlta " & _
    '                     "WHERE EstadoCreditoID IN (4,8,10) " & _
    '                     "AND OficinaCod IN (100, 107, 103, 112, 113, 114, 200, 201) " & _
    '                     "AND empleoid = " & [Forms]![002AltaCreditos]![EmpleoSolicitanteID]
                        
       
    ' Set RstCargosORI = OpenRS(DatoscargosORI)
    
    ' RunSQL ("DELETE * FROM TempGestionCobranza")
    
    ' If Not RstCargosORI.EOF Then
    '     Do Until RstCargosORI.EOF
    '         LiquidaBoletaCob RstCargosORI!CreditoID, 4, 100, 0, Now
    '         GraboComprobanteCargo True, RstCargosORI!CreditoID
    '         RstCargosORI.MoveNext
    '     Loop
    ' End If
    ' RstCargosORI.Close
    CalculoCargos([Forms]![002AltaCreditos]![EmpleoSolicitanteID], array(100, 107, 103, 112, 113, 114, 200, 201))

    datoscargos = "SELECT a.CreditoID, a.SucursalCod, a.CreditoCuenta, a.CreditoFechaProxVto, a.CreditoSdoCap, a.EstadoCreditoID, a.OficinaCod, a.empleoid, Sum(b.ImporteConcepto) AS SumaDeImporteConcepto, b.ComprobanteID " & _
                "FROM TempConsCredAlta AS a INNER JOIN TempGestionCobranza as b ON a.CreditoID = b.CreditoId " & _
                "GROUP BY a.CreditoID, a.SucursalCod, a.CreditoCuenta, a.CreditoFechaProxVto, a.CreditoSdoCap, a.EstadoCreditoID, a.OficinaCod, a.empleoid, b.ComprobanteID " & _
                "HAVING OficinaCod IN (100, 107, 103, 112, 113, 114, 200, 201)  AND a.empleoid = " & [Forms]![002AltaCreditos]![EmpleoSolicitanteID]

                
    
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

    Forms![002AltaCreditos].MuestroRetenciones.Requery
    Forms![002AltaCreditos].MuestroCargo.Requery
    End If
Exit Sub
ErrFrances:
    Mensaje Err.description
    log.error "Ocurrio un error en la rutina de liquidacion para la linea 204: {0} - {1}", Err.number, Err.description
End Sub
