import React, { useState, useEffect } from 'react';
import { supabase } from '../lib/supabase';
import { jsPDF } from "jspdf";
import autoTable from "jspdf-autotable";
import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';

export default function Nomina({ mostrarAlerta, listaInvernaderos }) {
  const [trabajadores, setTrabajadores] = useState([]);
  const [jornalesPendientes, setJornalesPendientes] = useState([]);
  const [valesPendientes, setValesPendientes] = useState([]);
  const [pagosHistoricos, setPagosHistoricos] = useState([]); 
  const [tabInterna, setTabInterna] = useState('planilla'); 
  const [cargando, setCargando] = useState(false);
  const [subTabLiquidacion, setSubTabLiquidacion] = useState('Sábado (Jornalero)');

  // Estados dinámicos para la Liquidación Central
  const [fechasCorteQuincena, setFechasCorteQuincena] = useState({});
  const [formasPagoModificadas, setFormasPagoModificadas] = useState({});
  const [cuentasModificadas, setCuentasModificadas] = useState({});
  const [invernaderosSeleccionados, setInvernaderosSeleccionados] = useState({}); // 👈 Invernadero asignado por pago

  // Formulario extendido de Trabajador
  const [formTrabajador, setFormTrabajador] = useState({ 
    id_editando: null, 
    nombre_completo: '', 
    cedula: '', 
    telefono: '', 
    pago_jornal_base: '', 
    tipo_pago: 'Sábado (Jornalero)',
    fecha_registro: new Date().toISOString().split('T')[0],
    forma_pago_predeterminada: 'Efectivo',
    numero_cuenta_predeterminado: ''
  });

  const [formJornal, setFormJornal] = useState({ id_editando: null, trabajador_id: '', fecha_labor: new Date().toISOString().split('T')[0], invernadero_id: '', tipo_labor: 'JORNAL', valor_pagar: '', observaciones: '' });
  const [formVale, setFormVale] = useState({ id_editando: null, trabajador_id: '', fecha_vale: new Date().toISOString().split('T')[0], monto_vale: '', motivo_nota: '' });

  const tiposPago = ["Sábado (Jornalero)", "Quincenal (Fijo)"];
  const formasPagoOpciones = ["Efectivo", "Bre-B (Pago Inmediato)", "Bancolombia Ahorros", "Bancolombia Corriente", "Nequi", "Daviplata", "Banco de Bogotá", "Davivienda", "Otro"];
  
  const formatoPesos = (valor) => new Intl.NumberFormat('es-CO', { style: 'currency', currency: 'COP', minimumFractionDigits: 0 }).format(valor || 0);

  useEffect(() => {
    cargarDatosNomina();
  }, []);

  const cargarDatosNomina = async () => {
    setCargando(true);
    try {
      const [resTrab, resJor, resVal, resPagos] = await Promise.all([
        supabase.from('nomina_trabajadores').select('*').eq('activo', true).order('nombre_completo', { ascending: true }),
        supabase.from('nomina_jornales').select('*, nomina_trabajadores(nombre_completo), invernaderos(nombre)').eq('liquidado', false).order('fecha_labor', { ascending: false }),
        supabase.from('nomina_vales').select('*, nomina_trabajadores(nombre_completo)').eq('descontado', false).order('fecha_vale', { ascending: false }),
        supabase.from('nomina_pagos_realizados').select('*').order('id', { ascending: false })
      ]);

      const dataTrabajadores = resTrab.data || [];
      const dataJornales = resJor.data || [];
      
      setTrabajadores(dataTrabajadores);
      setJornalesPendientes(dataJornales);
      setValesPendientes(resVal.data || []);
      setPagosHistoricos(resPagos.data || []);

      const fCorte = {};
      const fPago = {};
      const cNum = {};
      const invSel = {};
      
      dataTrabajadores.forEach(t => {
        fCorte[t.id] = new Date().toISOString().split('T')[0];
        fPago[t.id] = t.forma_pago_predeterminada || 'Efectivo';
        cNum[t.id] = t.numero_cuenta_predeterminado || '';

        // Detectar si el jornalero ya tiene asistencias cargadas en la semana
        const jornalesObrero = dataJornales.filter(j => j.trabajador_id === t.id);
        const invDelJornal = jornalesObrero.find(j => j.invernadero_id)?.invernadero_id;
        invSel[t.id] = invDelJornal || '';
      });
      
      setFechasCorteQuincena(fCorte);
      setFormasPagoModificadas(fPago);
      setCuentasModificadas(cNum);
      setInvernaderosSeleccionados(invSel);

    } catch (err) {
      console.error("Error en nómina:", err);
      mostrarAlerta("Error al sincronizar datos de nómina", "error");
    } finally {
      setCargando(false);
    }
  };

  const registrarTrabajador = async (e) => {
    e.preventDefault();
    const payload = {
      nombre_completo: formTrabajador.nombre_completo.toUpperCase().trim(),
      cedula: formTrabajador.cedula.trim(),
      telefono: formTrabajador.telefono.trim(),
      pago_jornal_base: parseFloat(formTrabajador.pago_jornal_base) || 0,
      tipo_pago: formTrabajador.tipo_pago,
      fecha_registro: formTrabajador.fecha_registro,
      forma_pago_predeterminada: formTrabajador.forma_pago_predeterminada,
      numero_cuenta_predeterminado: formTrabajador.numero_cuenta_predeterminado.trim()
    };

    try {
      let error;
      if (formTrabajador.id_editando) {
        error = (await supabase.from('nomina_trabajadores').update(payload).eq('id', formTrabajador.id_editando)).error;
      } else {
        error = (await supabase.from('nomina_trabajadores').insert([payload])).error;
      }
      if (error) throw error;
      mostrarAlerta(formTrabajador.id_editando ? "Trabajador actualizado con éxito" : "Trabajador registrado en directorio", "exito");
      limpiarFormTrabajador();
      await cargarDatosNomina();
    } catch (err) { mostrarAlerta("Error al procesar el colaborador", "error"); }
  };

  const prepararEdicionTrabajador = (item) => {
    setFormTrabajador({ 
      id_editando: item.id, 
      nombre_completo: item.nombre_completo, 
      cedula: item.cedula, 
      telefono: item.telefono || '', 
      pago_jornal_base: item.pago_jornal_base, 
      tipo_pago: item.tipo_pago || 'Sábado (Jornalero)',
      fecha_registro: item.fecha_registro || new Date().toISOString().split('T')[0],
      forma_pago_predeterminada: item.forma_pago_predeterminada || 'Efectivo',
      numero_cuenta_predeterminado: item.numero_cuenta_predeterminado || ''
    });
    setTabInterna('personal');
  };

  const eliminarTrabajadorLogico = async (id, nombre) => {
    if (window.confirm(`¿Estás seguro de inactivar a ${nombre}?`)) {
      const { error } = await supabase.from('nomina_trabajadores').update({ activo: false }).eq('id', id);
      if (error) mostrarAlerta("No se pudo inactivar", "error");
      else { mostrarAlerta("Trabajador inactivado con éxito"); await cargarDatosNomina(); }
    }
  };

  const limpiarFormTrabajador = () => setFormTrabajador({ id_editando: null, nombre_completo: '', cedula: '', telefono: '', pago_jornal_base: '', tipo_pago: 'Sábado (Jornalero)', fecha_registro: new Date().toISOString().split('T')[0], forma_pago_predeterminada: 'Efectivo', numero_cuenta_predeterminado: '' });
  
  const registrarJornalDiario = async (e) => {
    e.preventDefault();
    if (!formJornal.trabajador_id) return;
    let valorFinal = parseFloat(formJornal.valor_pagar) || trabajadores.find(t => t.id === formJornal.trabajador_id)?.pago_jornal_base || 0;
    const payload = { trabajador_id: formJornal.trabajador_id, fecha_labor: formJornal.fecha_labor, invernadero_id: formJornal.invernadero_id || null, tipo_labor: formJornal.tipo_labor.toUpperCase().trim(), valor_pagar: valorFinal, observaciones: formJornal.observaciones ? formJornal.observaciones.toUpperCase().trim() : null };
    try {
      const { error } = formJornal.id_editando ? await supabase.from('nomina_jornales').update(payload).eq('id', formJornal.id_editando) : await supabase.from('nomina_jornales').insert([payload]);
      if (error) throw error;
      mostrarAlerta("Asistencia procesada", "exito");
      setFormJornal({ id_editando: null, trabajador_id: '', fecha_labor: new Date().toISOString().split('T')[0], invernadero_id: '', tipo_labor: 'JORNAL', valor_pagar: '', observaciones: '' });
      await cargarDatosNomina();
    } catch (err) { mostrarAlerta("Error al registrar jornal", "error"); }
  };

  const registrarValeAdelanto = async (e) => {
    e.preventDefault();
    const monto = parseFloat(formVale.monto_vale);
    if (!formVale.trabajador_id || monto <= 0) return;
    try {
      const { error } = await supabase.from('nomina_vales').insert([{ trabajador_id: formVale.trabajador_id, fecha_vale: formVale.fecha_vale, monto_vale: monto, motivo_nota: formVale.motivo_nota.toUpperCase().trim() }]);
      if (error) throw error;
      mostrarAlerta("Vale autorizado", "exito");
      setFormVale({ id_editando: null, trabajador_id: '', fecha_vale: new Date().toISOString().split('T')[0], monto_vale: '', motivo_nota: '' });
      await cargarDatosNomina();
    } catch (err) { mostrarAlerta("Error al registrar el vale", "error"); }
  };

  const calcularPreLiquidacion = (filtroPago) => {
    const invernaderosActivosList = listaInvernaderos || [];

    return trabajadores
      .filter(t => (t.tipo_pago || 'Sábado (Jornalero)') === filtroPago)
      .map(t => {
        const jornalesObrero = jornalesPendientes.filter(j => j.trabajador_id === t.id);
        const valesObrero = valesPendientes.filter(v => v.trabajador_id === t.id);
        let totalGanado = filtroPago === 'Quincenal (Fijo)' ? t.pago_jornal_base : jornalesObrero.reduce((acc, j) => acc + parseFloat(j.valor_pagar || 0), 0);
        const totalVales = valesObrero.reduce((acc, v) => acc + parseFloat(v.monto_vale || 0), 0);
        
        // Determinar nombre del Invernadero desde los jornales cargados o la selección manual
        const invIdActual = invernaderosSeleccionados[t.id];
        const invObjeto = invernaderosActivosList.find(inv => String(inv.id) === String(invIdActual));
        
        const invNombresJornales = [...new Set(jornalesObrero.map(j => j.invernaderos?.nombre).filter(Boolean))];
        
        let invernaderoNombreFinal = 'GENERAL / VARIOS';
        if (invObjeto?.nombre) {
          invernaderoNombreFinal = invObjeto.nombre.toUpperCase();
        } else if (invNombresJornales.length > 0) {
          invernaderoNombreFinal = invNombresJornales.join(' / ').toUpperCase();
        }

        return { 
          id: t.id, 
          nombre: t.nombre_completo, 
          cedula: t.cedula || 'N/R', 
          telefono: t.telefono || 'N/R', 
          diasTrabajados: jornalesObrero.length, 
          totalGanado, 
          totalVales, 
          netoPagar: (totalGanado - totalVales), 
          tipo_pago: t.tipo_pago || 'Sábado (Jornalero)',
          invernadero_nombre: invernaderoNombreFinal,
          invernadero_id: invIdActual || jornalesObrero.find(j => j.invernadero_id)?.invernadero_id || null
        };
      })
      .filter(l => l.tipo_pago === 'Quincenal (Fijo)' || l.diasTrabajados > 0 || l.totalVales > 0);
  };

  const generarComprobanteNominaPDF = (item, fechaHistorica = null, fPagoHist = null, cHist = null, idPagoHistorico = null) => {
    const doc = new jsPDF({ orientation: 'portrait', unit: 'mm', format: [105, 148] });

    doc.setDrawColor(17, 112, 151); 
    doc.setLineWidth(0.8); 
    doc.rect(4, 4, 97, 140);

    try {
      doc.addImage('/Logopapel.png', 'PNG', 43.5, 6, 18, 18);
    } catch (e) {
      console.warn("Logo no disponible", e);
    }

    const fechaImpresion = fechaHistorica || fechasCorteQuincena[item.id] || new Date().toISOString().split('T')[0];
    const metodoMapeado = fPagoHist || formasPagoModificadas[item.id] || 'Efectivo';
    const numMapeado = cHist || cuentasModificadas[item.id] || '---';

    const idEntero = idPagoHistorico || item.pago_id;
    const codigoComprobante = idEntero ? `NOM-${String(idEntero).padStart(4, '0')}` : 'NOM-BORRADOR';

    doc.setFont("helvetica", "bold"); 
    doc.setFontSize(8); 
    doc.setTextColor(60, 60, 60);
    doc.text(`Comp. N°: ${codigoComprobante}`, 6, 11);

    doc.setFont("helvetica", "bold"); 
    doc.setFontSize(11); 
    doc.setTextColor(17, 112, 151);
    doc.text("COMPROBANTE DE PAGO", 52.5, 28, { align: "center" });

    const yBase = 32;
    doc.setFillColor(242, 242, 242); 
    doc.rect(6, yBase, 93, 5, 'F'); 
    doc.rect(6, yBase + 10, 93, 5, 'F');
    doc.setDrawColor(210, 210, 210); 
    doc.setLineWidth(0.2); 
    doc.rect(6, yBase, 93, 20);
    doc.line(6, yBase + 5, 99, yBase + 5); 
    doc.line(6, yBase + 10, 99, yBase + 10); 
    doc.line(6, yBase + 15, 99, yBase + 15); 
    doc.line(52, yBase, 52, yBase + 5);

    doc.setFont("helvetica", "bold"); doc.setFontSize(7); doc.setTextColor(0);
    doc.text("Fecha Periodo:", 7, yBase + 3.5); doc.setFont("helvetica", "normal"); doc.text(fechaImpresion, 30, yBase + 3.5);
    doc.setFont("helvetica", "bold"); doc.text("Colaborador:", 7, yBase + 8.5); doc.setFont("helvetica", "normal"); doc.text(item.nombre, 23, yBase + 8.5);
    doc.setFont("helvetica", "bold"); doc.text("Cédula / CC:", 7, yBase + 13.5); doc.setFont("helvetica", "normal"); doc.text(item.cedula, 21, yBase + 13.5);
    doc.setFont("helvetica", "bold"); doc.text("Forma Pago:", 7, yBase + 18.5); doc.setFont("helvetica", "normal"); doc.text(`${metodoMapeado} (${numMapeado})`, 22, yBase + 18.5);

    const esQuincenal = item.tipo_pago === 'Quincenal (Fijo)';
    const filasTabla = [[
      `${esQuincenal ? '1' : item.diasTrabajados || item.dias_respaldo || '---'}`,
      `${esQuincenal ? 'Quincena' : 'Días'}`,
      `${esQuincenal ? 'PAGO SALARIO FIJO QUINCENAL' : 'JORNALES LABORADOS ACUMULADOS'}`,
      "",
      `+${formatoPesos(item.totalGanado || item.monto_devengado)}`
    ]];

    if (item.totalVales > 0) {
      filasTabla.push(["1", "Global", "DESCUENTO DE VALES / ADELANTOS", "", `-${formatoPesos(item.totalVales)}`]);
    }

    autoTable(doc, {
      startY: yBase + 23, 
      margin: { left: 6, right: 6 },
      head: [["Cant.", "Unidad", "Detalle", "Precio", "Monto"]], 
      body: filasTabla, 
      theme: 'grid',
      styles: { font: 'helvetica', fontSize: 6.5, cellPadding: 1.5 }, 
      headStyles: { fillColor: [17, 112, 151], textColor: [255, 255, 255], halign: 'center', fontStyle: 'bold' }
    });

    const yTotal = doc.lastAutoTable.finalY + 6;
    doc.setFont("helvetica", "bold"); doc.setFontSize(8.5); doc.setTextColor(0);
    doc.text("VALOR ENTREGADO:", 43, yTotal); doc.text(formatoPesos(item.netoPagar || item.monto_pagado), 99, yTotal, { align: 'right' });

    const yFirmas = yTotal + 18;
    doc.setDrawColor(0); doc.setLineWidth(0.3);
    doc.line(8, yFirmas, 48, yFirmas); doc.line(56, yFirmas, 96, yFirmas);

    doc.setFont("helvetica", "bold"); doc.setFontSize(7); doc.setTextColor(0);
    doc.text("ENTREGUÉ CONFORME", 28, yFirmas + 4, { align: "center" });
    doc.text("RECIBÍ CONFORME", 76, yFirmas + 4, { align: "center" });

    doc.save(`RECIBO_NOMINA_${codigoComprobante}_${item.nombre.replace(/ /g, "_")}.pdf`);
  };

  const pagarNominaTrabajador = async (item) => {
    const fechaSeleccionada = fechasCorteQuincena[item.id] || new Date().toISOString().split('T')[0];
    const fPagoFinal = formasPagoModificadas[item.id] || 'Efectivo';
    const cNumFinal = cuentasModificadas[item.id] || '';
    const invNombreFinal = item.invernadero_nombre || 'GENERAL / VARIOS';

    if (!window.confirm(`¿Confirmas el pago definitivo (${invNombreFinal}) con método ${fPagoFinal} para ${item.nombre}?`)) return;

    try {
      const { data: pagoGuardado, error: errPago } = await supabase.from('nomina_pagos_realizados').insert([{
        trabajador_id: item.id,
        fecha_pago: fechaSeleccionada,
        monto_pagado: item.netoPagar, 
        monto_devengado: item.totalGanado, 
        monto_vales: item.totalVales, 
        periodo: fechaSeleccionada,
        dias_liquidados: item.diasTrabajados || 15,
        forma_pago_efectiva: fPagoFinal,
        numero_cuenta_efectivo: cNumFinal,
        invernadero_nombre: invNombreFinal
      }]).select('*').single();

      if (errPago) throw errPago;

      // Inyectar en bitácora vinculando el invernadero_id para el Dashboard de Mano de Obra
      await supabase.from('nomina_jornales').insert([{
        trabajador_id: item.id,
        invernadero_id: item.invernadero_id || null, // 👈 Se guarda la relación real con el invernadero
        fecha_labor: fechaSeleccionada,
        tipo_labor: item.tipo_pago === 'Quincenal (Fijo)' ? 'QUINCENA FIJA' : 'JORNAL SEMANAL TOTAL',
        valor_pagar: item.totalGanado,
        liquidado: true, 
        observaciones: `CIERRE CONSOLIDADO PERIODO ${fechaSeleccionada}. INVERNADERO: ${invNombreFinal}. PAGO: ${fPagoFinal.toUpperCase()}`
      }]);

      await supabase.from('nomina_vales').update({ descontado: true }).eq('trabajador_id', item.id).eq('descontado', false);
      await supabase.from('nomina_jornales').update({ liquidado: true }).eq('trabajador_id', item.id).eq('liquidado', false);

      mostrarAlerta(`Pago asentado exitosamente con Comp N°: NOM-${String(pagoGuardado.id).padStart(4, '0')}`, "exito");
      await cargarDatosNomina();

      generarComprobanteNominaPDF({ ...item, pago_id: pagoGuardado.id }, fechaSeleccionada, fPagoFinal, cNumFinal, pagoGuardado.id);

    } catch (err) { mostrarAlerta("Error al cerrar cuentas en el servidor", "error"); }
  };

  const resetearCicloCompleto = async () => {
    if (window.confirm("¿Está seguro de limpiar la planilla?")) {
      mostrarAlerta("Planilla diaria lista para el nuevo ciclo", "exito");
      await cargarDatosNomina();
    }
  };

  const exportarNominaAExcel = async () => {
    if (!pagosHistoricos || pagosHistoricos.length === 0) {
      mostrarAlerta("No hay historial de pagos de nómina para exportar", "error");
      return;
    }

    try {
      const workbook = new ExcelJS.Workbook();

      const pagosJornaleros = pagosHistoricos.filter(p => {
        const t = trabajadores.find(trab => trab.id === p.trabajador_id);
        return !t || (t.tipo_pago || 'Sábado (Jornalero)') === 'Sábado (Jornalero)';
      });

      const pagosFijos = pagosHistoricos.filter(p => {
        const t = trabajadores.find(trab => trab.id === p.trabajador_id);
        return t && t.tipo_pago === 'Quincenal (Fijo)';
      });

      const construirHojaNomina = (nombreHoja, tituloHeader, datos) => {
        const sheet = workbook.addWorksheet(nombreHoja);
        
        sheet.columns = [
          { header: 'COMP. N°', key: 'comp_num', width: 14 },
          { header: 'COLABORADOR / OPERARIO', key: 'nombre', width: 28 },
          { header: 'INVERNADERO', key: 'invernadero', width: 22 },
          { header: 'FECHA PERIODO', key: 'fecha', width: 16 },
          { header: 'CÉDULA / CC', key: 'cedula', width: 16 },
          { header: 'FORMA PAGO', key: 'forma_pago', width: 22 },
          { header: 'N° CUENTA / CELULAR', key: 'cuenta', width: 22 },
          { header: 'MONTO DEVENGADO (+)', key: 'devengado', width: 22 },
          { header: 'VALES DESCONTADOS (-)', key: 'vales', width: 22 },
          { header: 'NETO ENTREGADO (=)', key: 'neto', width: 22 }
        ];

        datos.forEach(p => {
          const t = trabajadores.find(trab => trab.id === p.trabajador_id);
          sheet.addRow({
            comp_num: `NOM-${String(p.id).padStart(4, '0')}`,
            nombre: (t?.nombre_completo || 'DESCONOCIDO').toUpperCase(),
            invernadero: (p.invernadero_nombre || 'GENERAL / VARIOS').toUpperCase(),
            fecha: p.fecha_pago || '',
            cedula: t?.cedula || 'N/R',
            forma_pago: (p.forma_pago_efectiva || 'Efectivo').toUpperCase(),
            cuenta: p.numero_cuenta_efectivo || '---',
            devengado: parseFloat(p.monto_devengado) || 0,
            vales: parseFloat(p.monto_vales) || 0,
            neto: parseFloat(p.monto_pagado) || 0
          });
        });

        const totalRowNumber = sheet.rowCount + 1;
        const ultimaFila = sheet.rowCount;

        const totalRow = sheet.addRow({
          cuenta: 'TOTALES ACUMULADOS:',
          devengado: { formula: `=SUM(H2:H${ultimaFila})` },
          vales: { formula: `=SUM(I2:I${ultimaFila})` },
          neto: { formula: `=SUM(J2:J${ultimaFila})` }
        });

        const headerRow = sheet.getRow(1);
        headerRow.height = 24;
        headerRow.eachCell(cell => {
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF117097' } };
          cell.font = { name: 'Arial', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
          cell.alignment = { vertical: 'middle', horizontal: 'center' };
        });

        sheet.eachRow((row, rowNumber) => {
          if (rowNumber === 1 || rowNumber === totalRowNumber) return;
          row.height = 20;
          const esCebra = rowNumber % 2 === 0;
          row.eachCell((cell, colNumber) => {
            cell.font = { name: 'Arial', size: 9 };
            if (esCebra) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEBF5FB' } };

            if ([1, 3, 4, 5, 6, 7].includes(colNumber)) cell.alignment = { vertical: 'middle', horizontal: 'center' };
            else if ([8, 9, 10].includes(colNumber)) cell.alignment = { vertical: 'middle', horizontal: 'right' };
            else cell.alignment = { vertical: 'middle', horizontal: 'left' };

            if ([8, 9, 10].includes(colNumber)) cell.numFmt = '"$"#,##0';
          });
        });

        totalRow.height = 22;
        totalRow.eachCell((cell, colNumber) => {
          cell.font = { name: 'Arial', size: 10, bold: true, color: { argb: 'FF0A4C68' } };
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD6EEFC' } };
          cell.border = {
            top: { style: 'thin', color: { argb: 'FF117097' } },
            bottom: { style: 'double', color: { argb: 'FF117097' } }
          };
          if (colNumber === 7) cell.alignment = { vertical: 'middle', horizontal: 'right' };
          if ([8, 9, 10].includes(colNumber)) {
            cell.alignment = { vertical: 'middle', horizontal: 'right' };
            cell.numFmt = '"$"#,##0';
          }
        });

        sheet.autoFilter = { from: { row: 1, column: 1 }, to: { row: ultimaFila, column: sheet.columnCount } };
      };

      construirHojaNomina('Jornaleros - Sábado', 'NÓMINA JORNALEROS', pagosJornaleros);
      construirHojaNomina('Fijos - Quincenal', 'NÓMINA SUELDOS FIJOS', pagosFijos);

      const buffer = await workbook.xlsx.writeBuffer();
      const fechaHoy = new Date().toISOString().split('T')[0];
      const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      saveAs(blob, `HISTORIAL_NOMINA_${fechaHoy}.xlsx`);

      mostrarAlerta("Historial de nómina exportado a Excel con Invernaderos", "exito");
    } catch (err) {
      console.error("Error al exportar Excel:", err);
      mostrarAlerta("Error al generar el archivo Excel", "error");
    }
  };

  const listaLiquidacionSabado = calcularPreLiquidacion('Sábado (Jornalero)');
  const listaLiquidacionQuincena = calcularPreLiquidacion('Quincenal (Fijo)');

  return (
    <div className="space-y-6 pb-20 text-slate-800">
      <div className="flex gap-2 border-b-2 border-gray-200 pb-2 overflow-x-auto">
        <button onClick={() => setTabInterna('planilla')} className={`px-4 py-2.5 rounded-xl font-black text-xs uppercase tracking-wider transition-all whitespace-nowrap ${tabInterna === 'planilla' ? 'bg-[#117097] text-white shadow-md' : 'bg-gray-200 text-gray-600 hover:bg-gray-300'}`}>📅 Planilla Diaria / Vales</button>
        <button onClick={() => setTabInterna('liquidacion')} className={`px-4 py-2.5 rounded-xl font-black text-xs uppercase tracking-wider transition-all whitespace-nowrap ${tabInterna === 'liquidacion' ? 'bg-[#117097] text-white shadow-md' : 'bg-gray-200 text-gray-600 hover:bg-gray-300'}`}>💰 Liquidación Central</button>
        <button onClick={() => setTabInterna('personal')} className={`px-4 py-2.5 rounded-xl font-black text-xs uppercase tracking-wider transition-all whitespace-nowrap ${tabInterna === 'personal' ? 'bg-[#117097] text-white shadow-md' : 'bg-gray-200 text-gray-600 hover:bg-gray-300'}`}>👥 Registro de Trabajadores</button>
      </div>

      {tabInterna === 'planilla' && (
        <div className="grid grid-cols-1 xl:grid-cols-2 gap-6">
          <div className="space-y-4">
            <div className="bg-white p-6 rounded-3xl shadow-xl border-t-8 border-[#117097]">
              <h3 className="font-black text-slate-800 uppercase text-xs italic mb-4">📝 Asistencia y Labores por Invernadero</h3>
              <form onSubmit={registrarJornalDiario} className="space-y-4">
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                  <div>
                    <label className="text-[10px] font-bold text-gray-400 uppercase px-1 italic">Fecha de Labor</label>
                    <input type="date" className="w-full border-2 p-2.5 rounded-xl font-bold text-sm outline-none focus:border-[#117097]" value={formJornal.fecha_labor} onChange={e => setFormJornal({...formJornal, fecha_labor: e.target.value})} required />
                  </div>
                  <div>
                    <label className="text-[10px] font-bold text-gray-400 uppercase px-1 italic">Invernadero</label>
                    <select className="w-full border-2 p-2.5 rounded-xl font-bold bg-white text-xs outline-none focus:border-[#117097]" value={formJornal.invernadero_id} onChange={e => setFormJornal({...formJornal, invernadero_id: e.target.value})}>
                      <option value="">GENERAL / TODA EL INVERNADERO</option>
                      {listaInvernaderos?.filter(inv => inv.activo !== false).map(inv => <option key={inv.id} value={inv.id}>{inv.nombre?.toUpperCase()}</option>)}
                    </select>
                  </div>
                </div>
                <div>
                  <label className="text-[10px] font-bold text-gray-400 uppercase px-1 italic">Seleccione Trabajador (Solo Jornaleros)</label>
                  <select className="w-full border-2 p-2.5 rounded-xl font-bold bg-white text-xs uppercase outline-none focus:border-[#117097]" value={formJornal.trabajador_id} onChange={e => setFormJornal({...formJornal, trabajador_id: e.target.value})} required>
                    <option value="">Seleccione operario...</option>
                    {trabajadores.filter(t => t.tipo_pago !== 'Quincenal (Fijo)').map(t => <option key={t.id} value={t.id}>{t.nombre_completo} (Tarifa Base: {formatoPesos(t.pago_jornal_base)})</option>)}
                  </select>
                </div>
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                  <div>
                    <label className="text-[10px] font-bold text-gray-400 uppercase px-1 italic">Tipo de Labor</label>
                    <input type="text" className="w-full border-2 p-2.5 rounded-xl font-bold text-xs uppercase outline-none focus:border-[#117097]" value={formJornal.tipo_labor} onChange={e => setFormJornal({...formJornal, tipo_labor: e.target.value})} placeholder="JORNAL / CONTRATO" required />
                  </div>
                  <div>
                    <label className="text-[10px] font-bold text-gray-400 uppercase px-1 italic">Tarifa Especial (COP)</label>
                    <input type="number" className="w-full border-2 p-2.5 rounded-xl font-bold text-xs outline-none focus:border-[#117097]" value={formJornal.valor_pagar} onChange={e => setFormJornal({...formJornal, valor_pagar: e.target.value})} placeholder="Vacío = Usa tarifa base" />
                  </div>
                </div>
                <div>
                  <label className="text-[10px] font-bold text-gray-400 uppercase px-1 italic">Detalle / Notas de Actividad</label>
                  <input type="text" className="w-full border-2 p-2.5 rounded-xl font-bold text-xs uppercase outline-none focus:border-[#117097]" value={formJornal.observaciones} onChange={e => setFormJornal({...formJornal, observaciones: e.target.value})} placeholder="Ej: Amarre de plantas" />
                </div>
                <button type="submit" className="w-full py-3.5 text-white font-black rounded-xl uppercase text-xs tracking-wider shadow-md bg-[#117097] hover:bg-[#0a4c68]">✓ Cargar Jornal en Planilla</button>
              </form>
            </div>

            <div className="bg-white rounded-2xl shadow-lg overflow-hidden border border-gray-200">
              <div className="bg-[#117097] p-2.5 text-white font-black text-[10px] uppercase tracking-wider pl-4">Planilla Acumulada de la Semana (Jornaleros)</div>
              <div className="max-h-56 overflow-y-auto text-[10px]">
                <table className="w-full text-left border-collapse">
                  <tbody className="divide-y font-bold text-slate-600">
                    {jornalesPendientes.length === 0 ? <tr><td className="p-4 text-center text-gray-400 italic">No hay jornales cargados esta semana</td></tr> : jornalesPendientes.map(j => (
                      <tr key={j.id} className="hover:bg-slate-50">
                        <td className="p-3 text-slate-900 uppercase truncate">{j.nomina_trabajadores?.nombre_completo}</td>
                        <td className="p-3 uppercase text-slate-400">{j.invernaderos?.nombre || 'General'}</td>
                        <td className="p-3 text-right text-emerald-700 font-black pr-4">{formatoPesos(j.valor_pagar)}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          </div>

          <div className="space-y-4">
            <div className="bg-white p-6 rounded-3xl shadow-xl border-t-8 border-[#117097]">
              <h3 className="font-black text-slate-800 uppercase text-xs italic mb-4">🎟️ Registrar Vales, Adelantos y Descuentos</h3>
              <form onSubmit={registrarValeAdelanto} className="space-y-4">
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                  <div>
                    <label className="text-[10px] font-bold text-gray-400 uppercase px-1 italic">Fecha Vale</label>
                    <input type="date" className="w-full border-2 p-2.5 rounded-xl font-bold text-sm outline-none focus:border-[#117097]" value={formVale.fecha_vale} onChange={e => setFormVale({...formVale, fecha_vale: e.target.value})} required />
                  </div>
                  <div>
                    <label className="text-[10px] font-bold text-gray-400 uppercase px-1 italic">Monto Adelantado (COP)</label>
                    <input type="number" className="w-full border-2 p-2.5 rounded-xl font-black text-emerald-700 text-sm outline-none focus:border-[#117097]" value={formVale.monto_vale} onChange={e => setFormVale({...formVale, monto_vale: e.target.value})} required />
                  </div>
                </div>
                <div>
                  <label className="text-[10px] font-bold text-gray-400 uppercase px-1 italic">Trabajador Beneficiario</label>
                  <select className="w-full border-2 p-2.5 rounded-xl font-bold bg-white text-xs uppercase outline-none focus:border-[#117097]" value={formVale.trabajador_id} onChange={e => setFormVale({...formVale, trabajador_id: e.target.value})} required>
                    <option value="">Seleccione operario...</option>
                    {trabajadores.map(t => <option key={t.id} value={t.id}>{t.nombre_completo} ({t.tipo_pago || 'Jornalero'})</option>)}
                  </select>
                </div>
                <div>
                  <label className="text-[10px] font-bold text-gray-400 uppercase px-1 italic">Motivo del Adelanto</label>
                  <input type="text" className="w-full border-2 p-2.5 rounded-xl font-bold text-xs uppercase outline-none focus:border-[#117097]" value={formVale.motivo_nota} onChange={e => setFormVale({...formVale, motivo_nota: e.target.value})} required />
                </div>
                <button type="submit" className="w-full py-3.5 text-white font-black rounded-xl uppercase text-xs tracking-wider shadow-md bg-[#117097] hover:bg-[#0a4c68]">💸 Registrar Adelanto / Descuento</button>
              </form>
            </div>
          </div>
        </div>
      )}

      {/* LIQUIDACIÓN CENTRAL */}
      {tabInterna === 'liquidacion' && (
        <div className="space-y-6">
          <div className="bg-white rounded-3xl shadow-xl overflow-hidden border border-gray-200">
            <div className="p-4 bg-[#117097] text-white font-black text-xs uppercase tracking-widest italic flex flex-col md:flex-row md:items-center justify-between gap-4">
              <div className="flex gap-2">
                <button onClick={() => setSubTabLiquidacion('Sábado (Jornalero)')} className={`px-3 py-1.5 rounded-lg text-[10px] font-black uppercase transition-all ${subTabLiquidacion === 'Sábado (Jornalero)' ? 'bg-[#0a4c68] text-white border border-white/20' : 'bg-[#117097] text-sky-100'}`}>🚜 Jornaleros (Sábado) ({listaLiquidacionSabado.length})</button>
                <button onClick={() => setSubTabLiquidacion('Quincenal (Fijo)')} className={`px-3 py-1.5 rounded-lg text-[10px] font-black uppercase transition-all ${subTabLiquidacion === 'Quincenal (Fijo)' ? 'bg-[#0a4c68] text-white border border-white/20' : 'bg-[#117097] text-sky-100'}`}>📅 Fijos (Quincenal) ({listaLiquidacionQuincena.length})</button>
              </div>
              <button onClick={resetearCicloCompleto} className="px-3 py-1.5 bg-[#0a4c68] hover:bg-black text-white text-[10px] font-black uppercase rounded-lg transition-colors border border-black/10">🧹 Limpiar Asistencia de Ciclo</button>
            </div>

            <div className="overflow-x-auto">
              <table className="w-full text-left border-collapse text-[11px]">
                <thead>
                  <tr className="bg-gray-300 text-slate-800 uppercase font-black border-b-2 border-gray-400">
                    <th className="p-4">Trabajador / Colaborador</th>
                    {/* ⚡ SECTOR DE ASIGNACIÓN DE INVERNADERO */}
                    <th className="p-4 text-center">Invernadero / Sede</th>
                    <th className="p-4 text-center">Definir Fecha Corte</th>
                    <th className="p-4 text-center w-3/12">Forma de Pago Efectiva / Cuenta</th>
                    <th className="p-4 text-right">Monto Devengado (+)</th>
                    <th className="p-4 text-right">Vales (-)</th>
                    <th className="p-4 text-right">Neto a Pagar (=)</th>
                    <th className="p-4 text-center">Acciones</th>
                  </tr>
                </thead>
                <tbody className="divide-y-2 divide-gray-400 font-bold text-slate-700">
                  {listaLiquidacionSabado.length === 0 && subTabLiquidacion === 'Sábado (Jornalero)' || listaLiquidacionQuincena.length === 0 && subTabLiquidacion === 'Quincenal (Fijo)' ? (
                    <tr>
                      <td colSpan="8" className="p-8 text-center text-gray-400 italic font-bold">No hay pagos pendientes de procesar en este periodo.</td>
                    </tr>
                  ) : (
                    (subTabLiquidacion === 'Sábado (Jornalero)' ? listaLiquidacionSabado : listaLiquidacionQuincena).map((item) => (
                      <tr key={item.id} className="bg-white hover:bg-sky-50 transition-colors">
                        <td className="p-4 font-black uppercase border-l-8 border-[#117097] text-slate-900">
                          {item.nombre}
                        </td>

                        {/* SELECTOR INTERACTIVO DE INVERNADERO EN TIEMPO REAL */}
                        <td className="p-4 text-center">
                          <select 
                            className="border-2 border-slate-200 p-1.5 rounded-xl text-[10px] font-black text-[#117097] bg-white outline-none focus:border-[#117097]"
                            value={invernaderosSeleccionados[item.id] || ''}
                            onChange={e => setInvernaderosSeleccionados({ ...invernaderosSeleccionados, [item.id]: e.target.value })}
                          >
                            <option value="">GENERAL / VARIOS</option>
                            {listaInvernaderos?.filter(inv => inv.activo !== false).map(inv => (
                              <option key={inv.id} value={inv.id}>{inv.nombre?.toUpperCase()}</option>
                            ))}
                          </select>
                        </td>

                        <td className="p-4 text-center">
                          <input type="date" className="border-2 p-1.5 rounded-xl text-center font-black text-xs text-[#117097] bg-amber-50/10 border-amber-100 outline-none focus:border-[#117097]" value={fechasCorteQuincena[item.id] || ''} onChange={e => setFechasCorteQuincena({ ...fechasCorteQuincena, [item.id]: e.target.value })} />
                        </td>
                        
                        <td className="p-4">
                          <div className="grid grid-cols-2 gap-1.5 max-w-[280px] mx-auto items-center">
                            <select 
                              className="w-full border-2 border-slate-200 p-1.5 rounded-xl text-[10px] font-black text-[#117097] bg-white outline-none focus:border-[#117097]" 
                              value={formasPagoModificadas[item.id] || 'Efectivo'} 
                              onChange={e => setFormasPagoModificadas({ ...formasPagoModificadas, [item.id]: e.target.value })}
                            >
                              {formasPagoOpciones.map(fp => <option key={fp} value={fp}>{fp}</option>)}
                            </select>
                            <input 
                              type="text" 
                              className="w-full border-2 border-dashed border-amber-300 p-1.5 rounded-xl bg-amber-50/40 text-[10px] text-center font-black text-[#117097] outline-none focus:border-[#117097]" 
                              value={cuentasModificadas[item.id] || ''} 
                              onChange={e => setCuentasModificadas({ ...cuentasModificadas, [item.id]: e.target.value })} 
                              placeholder="N° Cuenta / Celular" 
                            />
                          </div>
                        </td>

                        <td className="p-4 text-right font-black text-xs text-emerald-700">{formatoPesos(item.totalGanado)}</td>
                        <td className="p-4 text-right font-black text-xs text-red-600">{formatoPesos(item.totalVales)}</td>
                        <td className="p-4 text-right font-black text-sm bg-slate-50/50">{formatoPesos(item.netoPagar)}</td>
                        <td className="p-4">
                          <div className="flex justify-center gap-2">
                            <button type="button" onClick={() => generarComprobanteNominaPDF(item)} className="px-3 py-1.5 bg-[#0a4c68] text-white text-[10px] font-black uppercase rounded-lg hover:bg-black">🖨️ PDF</button>
                            <button type="button" onClick={() => pagarNominaTrabajador(item)} className="px-3 py-1.5 bg-emerald-700 text-white text-[10px] font-black uppercase rounded-lg hover:bg-emerald-800 shadow">💵 PAGAR</button>
                          </div>
                        </td>
                      </tr>
                    ))
                  )}
                </tbody>
              </table>
            </div>
          </div>

          {/* HISTORIAL GENERAL DE LIQUIDACIONES DESPLEGADO HASTA ABAJO */}
          <div className="bg-white rounded-3xl shadow-xl overflow-hidden border border-gray-200">
            <div className="p-3 bg-slate-800 text-white font-black text-[10px] uppercase tracking-wider px-4 flex justify-between items-center flex-wrap gap-2">
              <span>📚 Historial de Quincenas y Periodos Liquidados Anteriormente</span>
              <button
                onClick={exportarNominaAExcel}
                className="px-3 py-1 bg-emerald-600 text-white font-black rounded-lg shadow-md hover:bg-emerald-700 transition-colors flex items-center gap-1 text-[10px] uppercase tracking-wider"
              >
                📊 EXPORTAR A EXCEL HISTORIAL DE NÓMINA
              </button>
            </div>
            <div className="overflow-x-auto min-h-[300px]">
              <table className="w-full text-left border-collapse text-[10px]">
                <thead>
                  <tr className="bg-slate-100 text-slate-600 uppercase font-black border-b sticky top-0">
                    <th className="p-3 text-center">Comp. N°</th>
                    <th className="p-3">Trabajador / Colaborador</th>
                    <th className="p-3 text-center">Invernadero</th>
                    <th className="p-3 text-center">Fecha Periodo</th>
                    <th className="p-3 text-center">Forma Pago Efectiva</th>
                    <th className="p-3 text-right">Jornales (+)</th>
                    <th className="p-3 text-right">Vales (-)</th>
                    <th className="p-3 text-right">Neto Entregado (=)</th>
                    <th className="p-3 text-center">Acciones</th>
                  </tr>
                </thead>
                <tbody className="divide-y font-bold text-slate-600">
                  {pagosHistoricos.filter(p => trabajadores.find(t => t.id === p.trabajador_id)?.tipo_pago === subTabLiquidacion).map(p => {
                    const t = trabajadores.find(trab => trab.id === p.trabajador_id);
                    const compFormateado = `NOM-${String(p.id).padStart(4, '0')}`;
                    return (
                      <tr key={p.id} className="bg-emerald-50/10 hover:bg-emerald-50 transition-colors">
                        <td className="p-3 text-center font-black text-slate-500">{compFormateado}</td>
                        <td className="p-3 font-black text-slate-800 uppercase border-l-4 border-emerald-600">{t?.nombre_completo}</td>
                        <td className="p-3 text-center font-black text-[#117097] uppercase">{p.invernadero_nombre || 'GENERAL / VARIOS'}</td>
                        <td className="p-3 text-center font-black text-slate-600">{p.fecha_pago}</td>
                        <td className="p-3 text-center text-slate-500 font-bold uppercase">{p.forma_pago_efectiva || 'Efectivo'} {p.numero_cuenta_efectivo ? `(#${p.numero_cuenta_efectivo})` : ''}</td>
                        <td className="p-3 text-right text-slate-700">{formatoPesos(p.monto_devengado)}</td>
                        <td className="p-3 text-right text-red-500">-{formatoPesos(p.monto_vales)}</td>
                        <td className="p-3 text-right text-emerald-700 font-black">{formatoPesos(p.monto_pagado)}</td>
                        <td className="p-3 text-center">
                          <button type="button" onClick={() => generarComprobanteNominaPDF({ id: p.trabajador_id, nombre: t?.nombre_completo, cedula: t?.cedula, tipo_pago: subTabLiquidacion, netoPagar: p.monto_pagado, totalGanado: p.monto_devengado, totalVales: p.monto_vales, dias_respaldo: p.dias_liquidados }, p.fecha_pago, p.forma_pago_efectiva, p.numero_cuenta_efectivo, p.id)} className="px-2 py-1 bg-slate-700 text-white rounded hover:bg-black text-[9px] font-black">🖨️ PDF</button>
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      )}

      {/* REGISTRO DE TRABAJADORES AVANZADO */}
      {tabInterna === 'personal' && (
        <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
          <div className="bg-white p-6 rounded-3xl shadow-xl border-t-8 border-[#117097] h-fit">
            <h3 className="font-black text-slate-800 uppercase text-xs italic mb-5">👤 Ingresar Nuevo Trabajador</h3>
            <form onSubmit={registrarTrabajador} className="space-y-4">
              <div>
                <label className="text-[10px] font-bold text-gray-400 uppercase px-1 italic">Nombre Completo</label>
                <input type="text" className="w-full border-2 p-2.5 rounded-xl font-bold text-sm uppercase outline-none focus:border-[#117097]" value={formTrabajador.nombre_completo} onChange={e => setFormTrabajador({...formTrabajador, nombre_completo: e.target.value})} required />
              </div>
              <div className="grid grid-cols-2 gap-2">
                <div>
                  <label className="text-[10px] font-bold text-gray-400 uppercase px-1 italic">Cédula</label>
                  <input type="text" className="w-full border-2 p-2.5 rounded-xl font-bold text-xs outline-none focus:border-[#117097]" value={formTrabajador.cedula} onChange={e => setFormTrabajador({...formTrabajador, cedula: e.target.value})} required placeholder="Sin puntos" />
                </div>
                <div>
                  <label className="text-[10px] font-bold text-gray-400 uppercase px-1 italic">Teléfono</label>
                  <input type="text" className="w-full border-2 p-2.5 rounded-xl font-bold text-xs outline-none focus:border-[#117097]" value={formTrabajador.telefono} onChange={e => setFormTrabajador({...formTrabajador, telefono: e.target.value})} placeholder="Celular" />
                </div>
              </div>
              <div className="grid grid-cols-2 gap-2">
                <div>
                  <label className="text-[10px] font-bold text-gray-400 uppercase px-1 italic">Frecuencia de Pago</label>
                  <select className="w-full border-2 p-2.5 rounded-xl font-bold bg-white text-xs outline-none focus:border-[#117097]" value={formTrabajador.tipo_pago} onChange={e => setFormTrabajador({...formTrabajador, tipo_pago: e.target.value})}>
                    {tiposPago.map(p => <option key={p} value={p}>{p}</option>)}
                  </select>
                </div>
                <div>
                  <label className="text-[10px] font-bold text-gray-400 uppercase px-1 italic">Fecha de Registro / Ingreso</label>
                  <input type="date" className="w-full border-2 p-2.5 rounded-xl font-bold text-xs outline-none focus:border-[#117097]" value={formTrabajador.fecha_registro} onChange={e => setFormTrabajador({...formTrabajador, fecha_registro: e.target.value})} required />
                </div>
              </div>
              <div>
                <label className="text-[10px] font-bold text-gray-400 uppercase px-1 italic">{formTrabajador.tipo_pago === 'Quincenal (Fijo)' ? 'Sueldo Quincenal Fijo' : 'Tarifa Jornal Diario Base'}</label>
                <input type="number" className="w-full border-2 p-2.5 rounded-xl font-black text-sm text-slate-800 outline-none focus:border-[#117097]" value={formTrabajador.pago_jornal_base} onChange={e => setFormTrabajador({...formTrabajador, pago_jornal_base: e.target.value})} required placeholder="COP" />
              </div>
              <div className="bg-slate-50 p-3 rounded-2xl border-2 border-slate-100 space-y-2">
                <p className="text-[9px] font-black text-slate-500 uppercase tracking-tighter italic">Información de Pago / Transferencia Fija</p>
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-2">
                  <div>
                    <label className="text-[8px] font-bold text-slate-400 uppercase px-1">Forma / Banco</label>
                    <select className="w-full border p-2 bg-white rounded-xl text-xs font-bold outline-none focus:border-[#117097]" value={formTrabajador.forma_pago_predeterminada} onChange={e => setFormTrabajador({...formTrabajador, forma_pago_predeterminada: e.target.value})}>
                      {formasPagoOpciones.map(fp => <option key={fp} value={fp}>{fp}</option>)}
                    </select>
                  </div>
                  <div>
                    <label className="text-[8px] font-bold text-slate-400 uppercase px-1">N° Cuenta / Llave</label>
                    <input type="text" className="w-full border p-2 bg-white rounded-xl text-xs font-black text-[#117097] outline-none focus:border-[#117097]" value={formTrabajador.numero_cuenta_predeterminado} onChange={e => setFormTrabajador({...formTrabajador, numero_cuenta_predeterminado: e.target.value})} placeholder="Celular o cuenta..." />
                  </div>
                </div>
              </div>
              <div className="flex gap-2">
                {formTrabajador.id_editando && <button type="button" onClick={limpiarFormTrabajador} className="w-1/3 bg-gray-500 text-white font-black py-3 rounded-xl text-xs uppercase">X</button>}
                <button type="submit" className="flex-1 py-3.5 text-white font-black rounded-xl uppercase text-xs tracking-wider bg-[#117097] hover:bg-[#0a4c68] shadow-md">{formTrabajador.id_editando ? '💾 Actualizar' : '💾 Guardar Colaborador'}</button>
              </div>
            </form>
          </div>

          <div className="lg:col-span-2 bg-white rounded-3xl shadow-xl overflow-hidden border border-gray-200">
            <div className="p-4 bg-slate-800 text-white font-black text-xs flex justify-between items-center">
              <span>Directorio Interno de Operarios</span>
            </div>
            <div className="overflow-x-auto max-h-[520px] overflow-y-auto">
              <table className="w-full text-left border-collapse text-[11px]">
                <thead>
                  <tr className="bg-gray-200 text-slate-800 uppercase font-black border-b border-gray-300 sticky top-0">
                    <th className="p-3">Operario / Cédula</th>
                    <th className="p-3">Teléfono</th>
                    <th className="p-3">Forma de Pago / Cuenta</th>
                    <th className="p-3 text-center">Ingreso</th>
                    <th className="p-3 text-right">Frecuencia / Base</th>
                    <th className="p-3 text-center">Acciones</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-gray-200 font-bold text-slate-700">
                  {trabajadores.length === 0 ? (
                    <tr><td colSpan="6" className="p-6 text-center text-gray-400 italic">No hay operarios registrados</td></tr>
                  ) : (
                    trabajadores.map((t, idx) => (
                      <tr key={t.id} className={`${idx % 2 === 0 ? 'bg-white' : 'bg-slate-50'} hover:bg-sky-50 transition-colors`}>
                        <td className="p-3 uppercase">
                          <p className="font-black text-slate-900 text-xs">{t.nombre_completo}</p>
                          <p className="text-[10px] text-gray-400 font-bold">C.C. {t.cedula || 'N/R'}</p>
                        </td>
                        <td className="p-3">
                          <p className="text-slate-800 font-black">📞 {t.telefono || 'N/R'}</p>
                        </td>
                        <td className="p-3">
                          <span className="inline-block bg-amber-100 text-amber-900 text-[9px] px-2 py-0.5 rounded font-black uppercase">
                            💳 {t.forma_pago_predeterminada || 'Efectivo'}
                          </span>
                          {t.numero_cuenta_predeterminado && (
                            <p className="text-[9px] text-[#117097] font-black mt-0.5">#{t.numero_cuenta_predeterminado}</p>
                          )}
                        </td>
                        <td className="p-3 text-center">
                          <span className="text-[10px] text-slate-600 font-black">
                            🗓️ {t.fecha_registro || '---'}
                          </span>
                        </td>
                        <td className="p-3 text-right">
                          <span className="block text-[8px] uppercase font-black text-slate-400">
                            {t.tipo_pago === 'Quincenal (Fijo)' ? 'Sueldo Fijo' : 'Jornal Base'}
                          </span>
                          <span className="font-black text-slate-900 text-xs">
                            {formatoPesos(t.pago_jornal_base)}
                          </span>
                        </td>
                        <td className="p-3 text-center">
                          <div className="flex gap-1.5 justify-center">
                            <button onClick={() => prepararEdicionTrabajador(t)} className="p-1.5 bg-amber-100 text-amber-700 rounded-lg border border-amber-200 hover:bg-amber-600 hover:text-white transition-all text-xs" title="Editar">✏️</button>
                            <button onClick={() => eliminarTrabajadorLogico(t.id, t.nombre_completo)} className="p-1.5 bg-red-100 text-red-700 rounded-lg border border-red-200 hover:bg-red-600 hover:text-white transition-all text-xs" title="Inactivar">🗑️</button>
                          </div>
                        </td>
                      </tr>
                    ))
                  )}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}