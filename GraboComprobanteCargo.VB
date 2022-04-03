Public Sub GraboComprobanteCargo(Optional isCargoOtorgamiento As Boolean = False, Optional CreditoID As Long = 0)
    Dim RstComprobantes As Recordset
    Dim AuxcomprobID As Long
    
    On Error GoTo ErrComprob
    Set RstComprobantes = openrs("Comprobantes")

    If Not RstComprobantes.EOF Then
        RstComprobantes.AddNew
        RstComprobantes!UsuarioID = Forms!Menu!UsuarioID
        RstComprobantes!SucursalID = Forms!Menu!SucursalID
        RstComprobantes!ComprobanteFecha = Date
        RstComprobantes!ComprobanteFechaLiq = Date
        RstComprobantes!TipoComprobanteID = 9
        RstComprobantes!ComprobanteNumero = DameNumComprobante(RstComprobantes!TipoComprobanteID)
        RstComprobantes!ComprobanteImporte = DSum("ImporteConcepto", "TempGestionCobranza")
        RstComprobantes!PersonaId = Forms!Menu!PersonaId
        RstComprobantes!CreditoID = DMax("CreditoId", "TempGestionCobranza")
        RstComprobantes!ComprobanteTipo = 3
        RstComprobantes!ComprobanteEstado = "A"
        RstComprobantes.Update
    End If
    RstComprobantes.Close
    AuxcomprobID = DMax("ComprobanteID", "Comprobantes") 'AQUI HAY UN POTENCIAL ERROR POR CONCURRENCIA.

    RunSQL "INSERT INTO ComprobanteDetalles ( CreditoId, NumCuota, FechaVto, ConceptoCuotaID, ImporteConcepto, FechaLiquidacion, DatoAdicionalText, DatoAdicionalNum, ComprobanteID ) " & _
        "SELECT CreditoId, NumCuota, VtoCuota, ConceptoCuotaID, ImporteConcepto, FechaLiquidacion, DatoAdicionalText, DatoAdicionalNum, {0} AS ComprobanteID " & _
        "FROM TempGestionCobranza WHERE CreditoId = {1}", AuxcomprobID, CreditoID

    RunSQL "UPDATE TempGestionCobranza SET ComprobanteID = {0} WHERE ComprobanteID Is Null", AuxcomprobID 
    Exit Sub
ErrComprob:
    MsgBox "No Se Pudo Grabar Comprobante !!! error: " & Err.description, vbCritical, "Atención !!!"
End Sub