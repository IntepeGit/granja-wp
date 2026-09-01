import React, { useState, useEffect } from 'react';
import { supabase } from '../lib/supabase';
import { jsPDF } from "jspdf";
import autoTable from "jspdf-autotable";
import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';

export default function Nomina({ mostrarAlerta, listaInvernaderos }) {
  const obtenerFechaLocalHoy = () => {
    const d = new Date();
    const offset = d.getTimezoneOffset();
    const localDate = new Date(d.getTime() - (offset * 60 * 1000));
    return localDate.toISOString().split('T')[0];
  };

  const [trabajadores, setTrabajadores] = useState([]);
  const [jornalesPendientes, setJornalesPendientes] = useState([]);
  const [valesPendientes, setValesPendientes] = useState([]);
  const [pagosHistoricos, setPagosHistoricos] = useState([]); 
  const [tabInterna, setTabInterna] = useState('planilla'); 
  const [cargando, setCargando] = useState(false);
  const [subTabLiquidacion, setSubTabLiquidacion] = useState('Sábado (Jornalero)');

  const [modalActivo, setModalActivo] = useState(null); 
  const [pagoEditando, setPagoEditando] = useState(null);
  const [trabajadorDetalleEdicion, setTrabajadorDetalleEdicion] = useState(null);

  const [fechasCorteQuincena, setFechasCorteQuincena] = useState({});
  const [formasPagoModificadas, setFormasPagoModificadas] = useState({});
  const [cuentasModificadas, setCuentasModificadas] = useState({});
  const [invernaderosSeleccionados, setInvernaderosSeleccionados] = useState({});

  const [formTrabajador, setFormTrabajador] = useState({ 
    id_editando: null, nombre_completo: '', cedula: '', telefono: '', email: '',
    pago_jornal_base: '', tipo_pago: 'Sábado (Jornalero)', fecha_registro: obtenerFechaLocalHoy(),
    forma_pago_predeterminada: 'Efectivo', numero_cuenta_predeterminado: ''
  });

  const [formJornal, setFormJornal] = useState({ 
    id_editando: null, trabajador_id: '', fecha_labor: obtenerFechaLocalHoy(), invernadero_id: '', 
    tipo_labor: '', tarifa_hora: '', cantidad_horas: '', valor_pagar: '', observaciones: '' 
  });

  const [formVale, setFormVale] = useState({ 
    id_editando: null, trabajador_id: '', fecha_vale: obtenerFechaLocalHoy(), monto_vale: '', motivo_nota: '' 
  });

  const tiposPago = ["Sábado (Jornalero)", "Quincenal (Fijo)"];
  const formasPagoOpciones = ["Efectivo", "Bre-B (Pago Inmediato)", "Bancolombia Ahorros", "Bancolombia Corriente", "Nequi", "Daviplata", "Banco de Bogotá", "Davivienda", "Otro"];
  
  const formatoPesos = (valor) => new Intl.NumberFormat('es-CO', { style: 'currency', currency: 'COP', minimumFractionDigits: 0 }).format(valor || 0);

  useEffect(() => {
    cargarDatosNomina();
  }, []);

  const cargarDatosNomina = async () => {
    setCargando(true);
    try {
      const [resTrab, resJor, resVal, resPagos, resTodosJornales, resTodosVales] = await Promise.all([
        supabase.from('nomina_trabajadores').select('*').eq('activo', true).order('nombre_completo', { ascending: true }),
        supabase.from('nomina_jornales').select('*, nomina_trabajadores(nombre_completo, tipo_pago), invernaderos(nombre)').eq('liquidado', false).order('fecha_labor', { ascending: false }),
        supabase.from('nomina_vales').select('*, nomina_trabajadores(nombre_completo)').eq('descontado', false).order('fecha_vale', { ascending: false }),
        supabase.from('nomina_pagos_realizados').select('*').order('id', { ascending: false }),
        supabase.from('nomina_jornales').select('*, invernaderos(nombre)').order('id', { ascending: true }),
        supabase.from('nomina_vales').select('*').order('id', { ascending: true })
      ]);

      const dataTrabajadores = resTrab.data || [];
      const dataJornales = resJor.data || [];
      const todosJor = resTodosJornales.data || [];
      const todosVal = resTodosVales.data || [];

      const todosPagosOrdenados = [...(resPagos.data || [])].sort((a, b) => a.id - b.id);
      const jornalesAsignados = new Set();
      const valesAsignados = new Set();
      const mapaPagosConDetalle = {};

      todosPagosOrdenados.forEach(pago => {
        const tId = pago.trabajador_id;
        const targetDevengado = parseFloat(pago.monto_devengado) || 0;
        const targetVales = parseFloat(pago.monto_vales) || 0;

        let sumaAcumuladaDevengado = 0;
        const jDetalle = [];

        for (const j of todosJor) {
          if (jornalesAsignados.has(j.id)) continue;
          if (j.trabajador_id !== tId) continue;

          const obsStr = (j.observaciones || '').toUpperCase();
          const esRegistroCierre = j.tipo_labor.includes('SEMANAL') || j.tipo_labor.includes('QUINCENA') || obsStr.includes('CIERRE CONSOLIDADO');
          if (esRegistroCierre) continue;

          if (j.pago_id === pago.id || (!j.pago_id && targetDevengado > 0 && (sumaAcumuladaDevengado + parseFloat(j.valor_pagar || 0)) <= (targetDevengado + 1))) {
            jDetalle.push(j);
            jornalesAsignados.add(j.id);
            sumaAcumuladaDevengado += parseFloat(j.valor_pagar || 0);
          }
        }

        let sumaAcumuladaVales = 0;
        const vDetalle = [];

        for (const v of todosVal) {
          if (valesAsignados.has(v.id)) continue;
          if (v.trabajador_id !== tId) continue;

          if (v.pago_id === pago.id || (!v.pago_id && targetVales > 0 && (sumaAcumuladaVales + parseFloat(v.monto_vale || 0)) <= (targetVales + 1))) {
            vDetalle.push(v);
            valesAsignados.add(v.id);
            sumaAcumuladaVales += parseFloat(v.monto_vale || 0);
          }
        }

        mapaPagosConDetalle[pago.id] = { ...pago, jornalesDetalle: jDetalle, valesDetalle: vDetalle };
      });

      setTrabajadores(dataTrabajadores);
      setJornalesPendientes(dataJornales);
      setValesPendientes(resVal.data || []);
      setPagosHistoricos((resPagos.data || []).map(p => mapaPagosConDetalle[p.id] || p));

      const fCorte = {}; const fPago = {}; const cNum = {}; const invSel = {};
      dataTrabajadores.forEach(t => {
        fCorte[t.id] = obtenerFechaLocalHoy();
        fPago[t.id] = t.forma_pago_predeterminada || 'Efectivo';
        cNum[t.id] = t.numero_cuenta_predeterminado || '';
        const jornalesObrero = dataJornales.filter(j => j.trabajador_id === t.id);
        invSel[t.id] = jornalesObrero.find(j => j.invernadero_id)?.invernadero_id || '';
      });
      
      setFechasCorteQuincena(fCorte); setFormasPagoModificadas(fPago); setCuentasModificadas(cNum); setInvernaderosSeleccionados(invSel);
    } catch (err) {
      if (typeof mostrarAlerta === "function") mostrarAlerta("Error al sincronizar datos de nómina", "error");
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
      email: formTrabajador.email ? formTrabajador.email.toLowerCase().trim() : null,
      pago_jornal_base: parseFloat(formTrabajador.pago_jornal_base) || 0,
      tipo_pago: formTrabajador.tipo_pago,
      fecha_registro: formTrabajador.fecha_registro,
      forma_pago_predeterminada: formTrabajador.forma_pago_predeterminada,
      numero_cuenta_predeterminado: formTrabajador.numero_cuenta_predeterminado.trim()
    };

    try {
      const { error } = formTrabajador.id_editando 
        ? await supabase.from('nomina_trabajadores').update(payload).eq('id', formTrabajador.id_editando) 
        : await supabase.from('nomina_trabajadores').insert([payload]);
      
      if (error) throw error;
      if (typeof mostrarAlerta === "function") mostrarAlerta(formTrabajador.id_editando ? "Trabajador actualizado" : "Trabajador registrado", "exito");
      limpiarFormTrabajador();
      setModalActivo(null);
      await cargarDatosNomina();
    } catch (err) { 
      if (typeof mostrarAlerta === "function") mostrarAlerta("Error al procesar el colaborador", "error"); 
    }
  };

  const prepararEdicionTrabajador = (item) => {
    setFormTrabajador({ 
      id_editando: item.id, nombre_completo: item.nombre_completo, cedula: item.cedula, telefono: item.telefono || '', email: item.email || '',
      pago_jornal_base: item.pago_jornal_base, tipo_pago: item.tipo_pago || 'Sábado (Jornalero)', fecha_registro: item.fecha_registro || obtenerFechaLocalHoy(),
      forma_pago_predeterminada: item.forma_pago_predeterminada || 'Efectivo', numero_cuenta_predeterminado: item.numero_cuenta_predeterminado || ''
    });
    setModalActivo('trabajador');
  };

  const eliminarTrabajadorLogico = async (id, nombre) => {
    if (window.confirm(`¿Estás seguro de inactivar a ${nombre}?`)) {
      const { error } = await supabase.from('nomina_trabajadores').update({ activo: false }).eq('id', id);
      if (!error) { 
        if (typeof mostrarAlerta === "function") mostrarAlerta("Trabajador inactivado", "exito"); 
        await cargarDatosNomina(); 
      }
    }
  };

  const limpiarFormTrabajador = () => setFormTrabajador({ id_editando: null, nombre_completo: '', cedula: '', telefono: '', email: '', pago_jornal_base: '', tipo_pago: 'Sábado (Jornalero)', fecha_registro: obtenerFechaLocalHoy(), forma_pago_predeterminada: 'Efectivo', numero_cuenta_predeterminado: '' });
  
  const registrarJornalDiario = async (e) => {
    e.preventDefault();
    if (!formJornal.trabajador_id || !formJornal.tipo_labor) return;
    
    const trabSel = trabajadores.find(t => t.id === formJornal.trabajador_id);
    const esTipoHoras = formJornal.tipo_labor === 'HORAS';
    let valorFinal = 0;

    if (esTipoHoras) {
      const tHora = parseFloat(formJornal.tarifa_hora) || 0;
      const cantH = parseFloat(formJornal.cantidad_horas) || 0;
      valorFinal = formJornal.valor_pagar ? parseFloat(formJornal.valor_pagar) : (tHora * cantH);
    } else {
      valorFinal = formJornal.valor_pagar ? parseFloat(formJornal.valor_pagar) : parseFloat(trabSel?.pago_jornal_base || 0);
    }

    const detalleCalculo = esTipoHoras && formJornal.cantidad_horas 
      ? `[${formJornal.cantidad_horas} HRS @ ${formatoPesos(formJornal.tarifa_hora)}/H] ${formJornal.observaciones ? formJornal.observaciones.toUpperCase().trim() : ''}`
      : (formJornal.observaciones ? formJornal.observaciones.toUpperCase().trim() : null);

    const payload = { 
      trabajador_id: formJornal.trabajador_id, fecha_labor: formJornal.fecha_labor, invernadero_id: formJornal.invernadero_id || null, 
      tipo_labor: formJornal.tipo_labor.toUpperCase().trim(), valor_pagar: valorFinal, observaciones: detalleCalculo 
    };

    try {
      const { error } = formJornal.id_editando ? await supabase.from('nomina_jornales').update(payload).eq('id', formJornal.id_editando) : await supabase.from('nomina_jornales').insert([payload]);
      if (error) throw error;
      if (typeof mostrarAlerta === "function") mostrarAlerta(formJornal.id_editando ? "Registro actualizado" : "Labor cargada", "exito");
      setFormJornal({ id_editando: null, trabajador_id: '', fecha_labor: obtenerFechaLocalHoy(), invernadero_id: '', tipo_labor: '', tarifa_hora: '', cantidad_horas: '', valor_pagar: '', observaciones: '' });
      setModalActivo(null);
      await cargarDatosNomina();
    } catch (err) { 
      if (typeof mostrarAlerta === "function") mostrarAlerta("Error al registrar labor", "error"); 
    }
  };

  const prepararEdicionJornal = (jornal) => {
    setFormJornal({
      id_editando: jornal.id, trabajador_id: jornal.trabajador_id, fecha_labor: jornal.fecha_labor, invernadero_id: jornal.invernadero_id || '',
      tipo_labor: jornal.tipo_labor || 'JORNAL', tarifa_hora: '', cantidad_horas: '', valor_pagar: jornal.valor_pagar || '', observaciones: jornal.observaciones || ''
    });
    setModalActivo('labor');
  };

  const eliminarJornalPendiente = async (id) => {
    if (window.confirm("¿Desea eliminar este registro de la planilla?")) {
      const { error } = await supabase.from('nomina_jornales').delete().eq('id', id);
      if (!error) await cargarDatosNomina();
    }
  };

  const descartarConsolidadoTrabajador = async (trabajadorId, nombre) => {
    if (window.confirm(`¿Confirma eliminar TODOS los registros pendientes para ${nombre}?`)) {
      const { error } = await supabase.from('nomina_jornales').delete().eq('trabajador_id', trabajadorId).eq('liquidado', false);
      if (!error) await cargarDatosNomina();
    }
  };

  const registrarValeAdelanto = async (e) => {
    e.preventDefault();
    const monto = parseFloat(formVale.monto_vale);
    if (!formVale.trabajador_id || monto <= 0) return;

    const payload = { trabajador_id: formVale.trabajador_id, fecha_vale: formVale.fecha_vale, monto_vale: monto, motivo_nota: formVale.motivo_nota.toUpperCase().trim() };

    try {
      const { error } = formVale.id_editando ? await supabase.from('nomina_vales').update(payload).eq('id', formVale.id_editando) : await supabase.from('nomina_vales').insert([payload]);
      if (error) throw error;
      if (typeof mostrarAlerta === "function") mostrarAlerta(formVale.id_editando ? "Vale actualizado" : "Vale autorizado", "exito");
      setFormVale({ id_editando: null, trabajador_id: '', fecha_vale: obtenerFechaLocalHoy(), monto_vale: '', motivo_nota: '' });
      setModalActivo(null);
      await cargarDatosNomina();
    } catch (err) { 
      if (typeof mostrarAlerta === "function") mostrarAlerta("Error al registrar el vale", "error"); 
    }
  };

  const prepararEdicionVale = (vale) => {
    setFormVale({ id_editando: vale.id, trabajador_id: vale.trabajador_id, fecha_vale: vale.fecha_vale, monto_vale: vale.monto_vale || '', motivo_nota: vale.motivo_nota || '' });
    setModalActivo('vale');
  };

  const eliminarValePendiente = async (id) => {
    if (window.confirm("¿Desea eliminar este vale de adelanto?")) {
      const { error } = await supabase.from('nomina_vales').delete().eq('id', id);
      if (!error) await cargarDatosNomina();
    }
  };

  const calcularPreLiquidacion = (filtroPago) => {
    const invernaderosActivosList = listaInvernaderos || [];

    return trabajadores
      .filter(t => (t.tipo_pago || 'Sábado (Jornalero)') === filtroPago)
      .map(t => {
        const jornalesObrero = jornalesPendientes.filter(j => j.trabajador_id === t.id);
        const valesObrero = valesPendientes.filter(v => v.trabajador_id === t.id);

        const sumaPlanillaLabores = jornalesObrero.reduce((acc, j) => acc + parseFloat(j.valor_pagar || 0), 0);
        let totalGanado = filtroPago === 'Quincenal (Fijo)' ? (parseFloat(t.pago_jornal_base || 0) + sumaPlanillaLabores) : sumaPlanillaLabores;
        const totalVales = valesObrero.reduce((acc, v) => acc + parseFloat(v.monto_vale || 0), 0);
        const invIdActual = invernaderosSeleccionados[t.id];
        const invObjeto = invernaderosActivosList.find(inv => String(inv.id) === String(invIdActual));
        const invNombresJornales = [...new Set(jornalesObrero.map(j => j.invernaderos?.nombre).filter(Boolean))];
        
        let invernaderoNombreFinal = 'GENERAL / VARIOS';
        if (invObjeto?.nombre) invernaderoNombreFinal = invObjeto.nombre.toUpperCase();
        else if (invNombresJornales.length > 0) invernaderoNombreFinal = invNombresJornales.join(' / ').toUpperCase();

        return { 
          id: t.id, nombre: t.nombre_completo, cedula: t.cedula || 'N/R', telefono: t.telefono || 'N/R', diasTrabajados: jornalesObrero.length, 
          totalGanado, totalVales, netoPagar: (totalGanado - totalVales), tipo_pago: t.tipo_pago || 'Sábado (Jornalero)',
          invernadero_nombre: invernaderoNombreFinal, invernadero_id: invIdActual || jornalesObrero.find(j => j.invernadero_id)?.invernadero_id || null,
          jornalesDetalle: jornalesObrero, valesDetalle: valesObrero
        };
      })
      .filter(l => l.tipo_pago === 'Quincenal (Fijo)' || l.diasTrabajados > 0 || l.totalVales > 0);
  };

  const generarComprobanteNominaPDF = async (item, fechaHistorica = null, fPagoHist = null, cHist = null, idPagoHistorico = null) => {
    try {
      const doc = new jsPDF({ orientation: 'portrait', unit: 'mm', format: [105, 148] });

      doc.setDrawColor(17, 112, 151); doc.setLineWidth(0.8); doc.rect(4, 4, 97, 140);
      try { doc.addImage('/Logopapel.png', 'PNG', 43.5, 6, 18, 18); } catch (e) {}

      const fechaImpresion = fechaHistorica || fechasCorteQuincena[item.id] || obtenerFechaLocalHoy();
      const metodoMapeado = fPagoHist || formasPagoModificadas[item.id] || 'Efectivo';
      const numMapeado = cHist || cuentasModificadas[item.id] || '';

      const idEntero = idPagoHistorico || item.pago_id;
      const codigoComprobante = idEntero ? `NOM-${String(idEntero).padStart(4, '0')}` : 'NOM-BORRADOR';

      doc.setFont("helvetica", "bold"); doc.setFontSize(7.5); doc.setTextColor(80, 80, 80); doc.text(`VOUCHER N°: ${codigoComprobante}`, 6, 11);
      doc.setFont("helvetica", "bold"); doc.setFontSize(10.5); doc.setTextColor(17, 112, 151); doc.text("COMPROBANTE DE PAGO DE NÓMINA", 52.5, 28, { align: "center" });

      const yBase = 32;
      doc.setFillColor(245, 247, 250); doc.rect(6, yBase, 93, 5, 'F'); doc.rect(6, yBase + 10, 93, 5, 'F');
      doc.setDrawColor(210, 210, 210); doc.setLineWidth(0.2); doc.rect(6, yBase, 93, 20);
      doc.line(6, yBase + 5, 99, yBase + 5); doc.line(6, yBase + 10, 99, yBase + 10); doc.line(6, yBase + 15, 99, yBase + 15); doc.line(50, yBase + 10, 50, yBase + 15);

      doc.setFontSize(6.8); doc.setFont("helvetica", "bold"); doc.setTextColor(0);
      doc.text("FECHA PERIODO:", 8, yBase + 3.5); doc.setFont("helvetica", "normal"); doc.text(`${fechaImpresion}`, 34, yBase + 3.5);
      doc.setFont("helvetica", "bold"); doc.text("COLABORADOR:", 8, yBase + 8.5); doc.setFont("helvetica", "normal"); doc.text(`${(item.nombre || 'OPERARIO').toUpperCase()}`, 34, yBase + 8.5);
      doc.setFont("helvetica", "bold"); doc.text("CÉDULA / CC:", 8, yBase + 13.5); doc.setFont("helvetica", "normal"); doc.text(`${item.cedula || 'N/R'}`, 28, yBase + 13.5);
      doc.setFont("helvetica", "bold"); doc.text("INVERNADERO:", 52, yBase + 13.5); doc.setFont("helvetica", "normal"); doc.text(`${(item.invernadero_nombre || 'GENERAL').toUpperCase()}`, 72, yBase + 13.5);
      doc.setFont("helvetica", "bold"); doc.text("FORMA DE PAGO:", 8, yBase + 18.5); doc.setFont("helvetica", "normal");
      
      const textoCuenta = numMapeado && numMapeado !== '---' ? ` (#${numMapeado})` : '';
      doc.text(`${metodoMapeado.toUpperCase()}${textoCuenta}`, 34, yBase + 18.5);

      const esQuincenal = item.tipo_pago === 'Quincenal (Fijo)';
      const filasTabla = [];

      if (esQuincenal) {
        const trabObj = trabajadores.find(t => t.id === item.id);
        const sueldoBase = parseFloat(trabObj?.pago_jornal_base || item.totalGanado || 0);
        filasTabla.push(["1", "Quincena", "PAGO SALARIO FIJO QUINCENAL BASE", `+${formatoPesos(sueldoBase)}`]);

        if (item.jornalesDetalle && item.jornalesDetalle.length > 0) {
          item.jornalesDetalle.forEach(j => {
            const unidadStr = j.tipo_labor === 'HORAS' ? 'Horas' : 'Jornal';
            const fechaLab = j.fecha_labor ? `[${j.fecha_labor}] ` : '';
            filasTabla.push(["1", unidadStr, `EXTRA: ${fechaLab}${j.tipo_labor}${j.observaciones ? ' - ' + j.observaciones : ''}`.toUpperCase(), `+${formatoPesos(j.valor_pagar)}`]);
          });
        }
      } else if (item.jornalesDetalle && item.jornalesDetalle.length > 0) {
        item.jornalesDetalle.forEach(j => {
          const unidadStr = j.tipo_labor === 'HORAS' ? 'Horas' : 'Jornal';
          const fechaLab = j.fecha_labor ? `[${j.fecha_labor}] ` : '';
          filasTabla.push(["1", unidadStr, `${fechaLab}${j.tipo_labor}${j.observaciones ? ' - ' + j.observaciones : ''}`.toUpperCase(), `+${formatoPesos(j.valor_pagar)}`]);
        });
      } else {
        filasTabla.push([`${item.diasTrabajados || item.dias_respaldo || '1'}`, "Jornal(es)", "JORNALES / LABORES ACUMULADAS PERIODO", `+${formatoPesos(item.totalGanado || item.monto_devengado)}`]);
      }

      if (item.valesDetalle && item.valesDetalle.length > 0) {
        item.valesDetalle.forEach(v => {
          const fechaV = v.fecha_vale ? `[${v.fecha_vale}] ` : '';
          filasTabla.push(["1", "Descuento", `${fechaV}VALE / ADELANTO: ${v.motivo_nota || 'PRESTAMO'}`, `-${formatoPesos(v.monto_vale)}`]);
        });
      } else if (item.totalVales > 0 || item.monto_vales > 0) {
        const valesMonto = item.totalVales || item.monto_vales;
        filasTabla.push(["1", "Descuento", "DESCUENTO DE VALES / ADELANTOS", `-${formatoPesos(valesMonto)}`]);
      }

      autoTable(doc, {
        startY: yBase + 23, margin: { left: 6, right: 6 },
        head: [["CANT.", "UNIDAD", "FECHA - CONCEPTO / LABOR Y NOTA", "MONTO"]], body: filasTabla, theme: 'grid',
        styles: { font: 'helvetica', fontSize: 6, cellPadding: 1.2, lineWidth: 0.1, lineColor: [210, 210, 210] }, 
        headStyles: { fillColor: [17, 112, 151], textColor: [255, 255, 255], halign: 'center', fontStyle: 'bold' },
        columnStyles: { 0: { cellWidth: 10, halign: 'center' }, 1: { cellWidth: 16, halign: 'center' }, 2: { cellWidth: 47, halign: 'left' }, 3: { cellWidth: 20, halign: 'right' } }
      });

      const yTotal = doc.lastAutoTable.finalY + 4;
      doc.setFillColor(235, 245, 238); doc.rect(48, yTotal, 51, 12, 'F');
      doc.setDrawColor(180, 210, 190); doc.rect(48, yTotal, 51, 12);
      doc.setFont("helvetica", "bold"); doc.setFontSize(7); doc.setTextColor(20, 90, 40); doc.text("TOTAL NETO ENTREGADO:", 73.5, yTotal + 4, { align: 'center' });
      doc.setFont("helvetica", "bold"); doc.setFontSize(10); doc.setTextColor(10, 80, 30); doc.text(formatoPesos(item.netoPagar || item.monto_pagado), 73.5, yTotal + 9.5, { align: 'center' });

      const yFirmas = yTotal + 22;
      doc.setDrawColor(120, 120, 120); doc.setLineWidth(0.3); doc.line(8, yFirmas, 48, yFirmas); doc.line(56, yFirmas, 96, yFirmas);
      doc.setFont("helvetica", "bold"); doc.setFontSize(6.5); doc.setTextColor(50);
      doc.text("FIRMA / CONFORME TRABAJADOR", 28, yFirmas + 3.5, { align: "center" }); doc.setFont("helvetica", "normal"); doc.text(`C.C. ${item.cedula || '----------------'}`, 28, yFirmas + 6.5, { align: "center" });
      doc.setFont("helvetica", "bold"); doc.text("ADMINISTRACIÓN / GRANJA", 76, yFirmas + 3.5, { align: "center" }); doc.setFont("helvetica", "normal"); doc.text("ENTREGADO Y REVISADO", 76, yFirmas + 6.5, { align: "center" });

      doc.save(`VOUCHER_NOMINA_${codigoComprobante}_${item.nombre.replace(/\s+/g, "_")}.pdf`);
    } catch (err) {
      if (typeof mostrarAlerta === "function") mostrarAlerta("Error al generar el comprobante PDF", "error");
    }
  };

  const pagarNominaTrabajador = async (item) => {
    const fechaSeleccionada = fechasCorteQuincena[item.id] || obtenerFechaLocalHoy();
    const fPagoFinal = formasPagoModificadas[item.id] || 'Efectivo';
    const cNumFinal = cuentasModificadas[item.id] || '';
    const invNombreFinal = item.invernadero_nombre || 'GENERAL / VARIOS';

    if (!window.confirm(`¿Confirmas el pago definitivo (${invNombreFinal}) con método ${fPagoFinal} para ${item.nombre}?`)) return;

    try {
      const { data: pagoGuardado, error: errPago } = await supabase.from('nomina_pagos_realizados').insert([{
        trabajador_id: item.id, fecha_pago: fechaSeleccionada, monto_pagado: item.netoPagar, monto_devengado: item.totalGanado, 
        monto_vales: item.totalVales, periodo: fechaSeleccionada, dias_liquidados: item.diasTrabajados || 15,
        forma_pago_efectiva: fPagoFinal, numero_cuenta_efectivo: cNumFinal, invernadero_nombre: invNombreFinal
      }]).select('*').single();

      if (errPago) throw errPago;
      const idPago = pagoGuardado.id;

      await supabase.from('nomina_jornales').update({ liquidado: true, pago_id: idPago }).eq('trabajador_id', item.id).eq('liquidado', false);
      await supabase.from('nomina_vales').update({ descontado: true, pago_id: idPago }).eq('trabajador_id', item.id).eq('descontado', false);

      if (typeof mostrarAlerta === "function") mostrarAlerta(`Pago asentado exitosamente con Comp N°: NOM-${String(idPago).padStart(4, '0')}`, "exito");
      await cargarDatosNomina();
      await generarComprobanteNominaPDF({ ...item, pago_id: idPago }, fechaSeleccionada, fPagoFinal, cNumFinal, idPago);
    } catch (err) { 
      if (typeof mostrarAlerta === "function") mostrarAlerta("Error al cerrar cuentas en el servidor", "error"); 
    }
  };

  const guardarEdicionPagoHistorico = async () => {
    if (!pagoEditando) return;
    try {
      const devengado = parseFloat(pagoEditando.monto_devengado) || 0;
      const vales = parseFloat(pagoEditando.monto_vales) || 0;
      const { error } = await supabase.from('nomina_pagos_realizados').update({
        fecha_pago: pagoEditando.fecha_pago, invernadero_nombre: pagoEditando.invernadero_nombre, forma_pago_efectiva: pagoEditando.forma_pago_efectiva,
        numero_cuenta_efectivo: pagoEditando.numero_cuenta_efectivo, monto_devengado: devengado, monto_vales: vales, monto_pagado: devengado - vales
      }).eq('id', pagoEditando.id);
      if (error) throw error;

      if (typeof mostrarAlerta === "function") mostrarAlerta("Registro de pago actualizado", "exito");
      setPagoEditando(null);
      await cargarDatosNomina();
    } catch (err) {
      if (typeof mostrarAlerta === "function") mostrarAlerta("Error al actualizar el pago", "error");
    }
  };

  const eliminarPagoHistorico = async (pago) => {
    if (window.confirm(`¿Confirma eliminar definitivamente el pago de nómina NOM-${String(pago.id).padStart(4, '0')}?`)) {
      const { error } = await supabase.from('nomina_pagos_realizados').delete().eq('id', pago.id);
      if (!error) await cargarDatosNomina();
    }
  };

  const resetearCicloCompleto = async () => {
    if (window.confirm("¿Está seguro de limpiar la asistencia?")) {
      if (typeof mostrarAlerta === "function") mostrarAlerta("Planilla diaria lista", "exito");
      await cargarDatosNomina();
    }
  };

  const exportarNominaAExcel = async () => {
    if (!pagosHistoricos || pagosHistoricos.length === 0) {
      if (typeof mostrarAlerta === "function") mostrarAlerta("No hay historial de pagos de nómina para exportar", "error");
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

      const construirHojaNominaDetallada = (nombreHoja, tituloHeader, datos) => {
        const sheet = workbook.addWorksheet(nombreHoja);
        sheet.columns = [
          { header: 'COMP. N°', key: 'comp_num', width: 14 },
          { header: 'COLABORADOR / OPERARIO', key: 'nombre', width: 28 },
          { header: 'INVERNADERO', key: 'invernadero', width: 22 },
          { header: 'FECHA PAGO / PERIODO', key: 'fecha_pago', width: 20 },
          { header: 'FECHA LABOR / VALE', key: 'fecha_item', width: 18 },
          { header: 'CONCEPTO / TIPO LABOR', key: 'concepto', width: 25 },
          { header: 'DETALLE / NOTAS / HORAS', key: 'detalle', width: 35 },
          { header: 'DEVENGADO (+)', key: 'devengado', width: 18 },
          { header: 'VALES / DESCUENTOS (-)', key: 'vales', width: 22 },
          { header: 'NETO ENTREGADO (=)', key: 'neto', width: 20 },
          { header: 'TOTAL VOUCHER (COP)', key: 'total_voucher', width: 22 }
        ];

        datos.forEach(p => {
          const t = trabajadores.find(trab => trab.id === p.trabajador_id);
          const compN = `NOM-${String(p.id).padStart(4, '0')}`;
          const nombreOp = (t?.nombre_completo || 'DESCONOCIDO').toUpperCase();
          const invOp = (p.invernadero_nombre || 'GENERAL / VARIOS').toUpperCase();
          const totalNetoVoucher = parseFloat(p.monto_pagado) || 0;
          const esFijo = t?.tipo_pago === 'Quincenal (Fijo)';

          const filasDelComprobante = [];

          if (esFijo) {
            const sueldoBase = parseFloat(t?.pago_jornal_base || p.monto_devengado || 0);
            filasDelComprobante.push({ comp_num: compN, nombre: nombreOp, invernadero: invOp, fecha_pago: p.fecha_pago || '', fecha_item: p.fecha_pago || '', concepto: 'SALARIO FIJO QUINCENAL', detalle: 'SUELDO BASE QUINCENAL', devengado: sueldoBase, vales: 0, neto: sueldoBase, total_voucher: null });

            if (p.jornalesDetalle && p.jornalesDetalle.length > 0) {
              p.jornalesDetalle.forEach(j => {
                filasDelComprobante.push({ comp_num: compN, nombre: nombreOp, invernadero: (j.invernaderos?.nombre || invOp).toUpperCase(), fecha_pago: p.fecha_pago || '', fecha_item: j.fecha_labor || '', concepto: `EXTRA: ${(j.tipo_labor || 'JORNAL').toUpperCase()}`, detalle: (j.observaciones || 'LABOR AGRÍCOLA EXTRA').toUpperCase(), devengado: parseFloat(j.valor_pagar) || 0, vales: 0, neto: parseFloat(j.valor_pagar) || 0, total_voucher: null });
              });
            }
          } else if (p.jornalesDetalle && p.jornalesDetalle.length > 0) {
            p.jornalesDetalle.forEach(j => {
              filasDelComprobante.push({ comp_num: compN, nombre: nombreOp, invernadero: (j.invernaderos?.nombre || invOp).toUpperCase(), fecha_pago: p.fecha_pago || '', fecha_item: j.fecha_labor || '', concepto: (j.tipo_labor || 'JORNAL').toUpperCase(), detalle: (j.observaciones || 'LABOR AGRÍCOLA').toUpperCase(), devengado: parseFloat(j.valor_pagar) || 0, vales: 0, neto: parseFloat(j.valor_pagar) || 0, total_voucher: null });
            });
          } else {
            filasDelComprobante.push({ comp_num: compN, nombre: nombreOp, invernadero: invOp, fecha_pago: p.fecha_pago || '', fecha_item: p.fecha_pago || '', concepto: 'PAGO CONSOLIDADO', detalle: 'TOTAL DEVENGADO DEL PERIODO', devengado: parseFloat(p.monto_devengado) || 0, vales: 0, neto: parseFloat(p.monto_devengado) || 0, total_voucher: null });
          }

          if (p.valesDetalle && p.valesDetalle.length > 0) {
            p.valesDetalle.forEach(v => {
              filasDelComprobante.push({ comp_num: compN, nombre: nombreOp, invernadero: invOp, fecha_pago: p.fecha_pago || '', fecha_item: v.fecha_vale || '', concepto: 'DESCUENTO DE VALE', detalle: (v.motivo_nota || 'PRESTAMO / ADELANTO').toUpperCase(), devengado: 0, vales: parseFloat(v.monto_vale) || 0, neto: -(parseFloat(v.monto_vale) || 0), total_voucher: null });
            });
          } else if (parseFloat(p.monto_vales) > 0) {
            filasDelComprobante.push({ comp_num: compN, nombre: nombreOp, invernadero: invOp, fecha_pago: p.fecha_pago || '', fecha_item: p.fecha_pago || '', concepto: 'DESCUENTO GLOBAL', detalle: 'VALES / ADELANTOS ACUMULADOS', devengado: 0, vales: parseFloat(p.monto_vales) || 0, neto: -(parseFloat(p.monto_vales) || 0), total_voucher: null });
          }

          if (filasDelComprobante.length > 0) {
            filasDelComprobante[filasDelComprobante.length - 1].total_voucher = totalNetoVoucher;
          }

          filasDelComprobante.forEach(filaData => {
            const addedRow = sheet.addRow(filaData);
            if (filaData.total_voucher !== null) {
              const celdaTotal = addedRow.getCell(11);
              celdaTotal.font = { name: 'Arial', size: 9, bold: true, color: { argb: 'FF0A4C68' } };
              celdaTotal.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEBF5FB' } };
              celdaTotal.border = { top: { style: 'thin', color: { argb: 'FF117097' } }, bottom: { style: 'double', color: { argb: 'FF117097' } } };
            }
          });
        });

        const totalRowNumber = sheet.rowCount + 1;
        const ultimaFila = sheet.rowCount;

        const totalRow = sheet.addRow({
          detalle: 'TOTALES ACUMULADOS:',
          devengado: { formula: `=SUM(H2:H${ultimaFila})` },
          vales: { formula: `=SUM(I2:I${ultimaFila})` },
          neto: { formula: `=SUM(J2:J${ultimaFila})` },
          total_voucher: { formula: `=SUM(K2:K${ultimaFila})` }
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
            if (esCebra && colNumber !== 11) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEBF5FB' } };

            if ([1, 3, 4, 5, 6].includes(colNumber)) cell.alignment = { vertical: 'middle', horizontal: 'center' };
            else if ([8, 9, 10, 11].includes(colNumber)) cell.alignment = { vertical: 'middle', horizontal: 'right' };
            else cell.alignment = { vertical: 'middle', horizontal: 'left' };

            if ([8, 9, 10, 11].includes(colNumber)) cell.numFmt = '"$"#,##0';
          });
        });

        totalRow.height = 22;
        totalRow.eachCell((cell, colNumber) => {
          cell.font = { name: 'Arial', size: 10, bold: true, color: { argb: 'FF0A4C68' } };
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD6EEFC' } };
          cell.border = { top: { style: 'thin', color: { argb: 'FF117097' } }, bottom: { style: 'double', color: { argb: 'FF117097' } } };
          if (colNumber === 7) cell.alignment = { vertical: 'middle', horizontal: 'right' };
          if ([8, 9, 10, 11].includes(colNumber)) { cell.alignment = { vertical: 'middle', horizontal: 'right' }; cell.numFmt = '"$"#,##0'; }
        });
        sheet.autoFilter = { from: { row: 1, column: 1 }, to: { row: ultimaFila, column: sheet.columnCount } };
      };

      construirHojaNominaDetallada('Jornaleros - Sábado', 'NÓMINA JORNALEROS', pagosJornaleros);
      construirHojaNominaDetallada('Fijos - Quincenal', 'NÓMINA SUELDOS FIJOS', pagosFijos);

      const buffer = await workbook.xlsx.writeBuffer();
      const fechaHoy = obtenerFechaLocalHoy();
      const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      saveAs(blob, `HISTORIAL_NOMINA_DESGLOSADO_${fechaHoy}.xlsx`);
    } catch (err) {
      if (typeof mostrarAlerta === "function") mostrarAlerta("Error al generar el archivo Excel", "error");
    }
  };

  const listaLiquidacionSabado = calcularPreLiquidacion('Sábado (Jornalero)');
  const listaLiquidacionQuincena = calcularPreLiquidacion('Quincenal (Fijo)');

  const trabSeleccionadoObjeto = trabajadores.find(t => t.id === formJornal.trabajador_id);
  const esTrabajadorQuincenal = trabSeleccionadoObjeto?.tipo_pago === 'Quincenal (Fijo)';
  const esLaborPorHoras = formJornal.tipo_labor === 'HORAS';

  const listaPlanillaUnificada = [
    ...jornalesPendientes.map(j => ({
      id: j.id, tipo_registro: 'LABOR', fecha: j.fecha_labor, trabajador_nombre: j.nomina_trabajadores?.nombre_completo || 'DESCONOCIDO',
      invernadero_nombre: j.invernaderos?.nombre || 'GENERAL / VARIOS', concepto_tipo: j.tipo_labor, detalle_nota: j.observaciones || 'SIN NOTAS',
      monto_devengado: j.valor_pagar, monto_descuento: 0, objeto_original: j
    })),
    ...valesPendientes.map(v => ({
      id: v.id, tipo_registro: 'VALE', fecha: v.fecha_vale, trabajador_nombre: v.nomina_trabajadores?.nombre_completo || 'DESCONOCIDO',
      invernadero_nombre: 'GENERAL (VALE)', concepto_tipo: '🎟️ VALE / ADELANTO', detalle_nota: v.motivo_nota || 'SIN MOTIVO',
      monto_devengado: 0, monto_descuento: v.monto_vale, objeto_original: v
    }))
  ].sort((a, b) => new Date(b.fecha) - new Date(a.fecha));

  return (
    <div className="space-y-6 pb-20 text-slate-800 dark:text-slate-200 font-sans transition-colors duration-300">
      
      {/* 🚀 CABECERA PRINCIPAL MODERNA */}
      <div className="bg-white dark:bg-slate-800 p-6 rounded-3xl shadow-xl border border-slate-200 dark:border-slate-700 flex flex-col xl:flex-row justify-between items-center gap-4 transition-colors duration-300">
        <div className="space-y-1 w-full xl:w-auto text-center xl:text-left">
          <div className="flex items-center justify-center xl:justify-start gap-3">
            <span className="text-2xl p-2 bg-[#117097]/10 dark:bg-sky-500/20 rounded-xl text-[#117097] dark:text-sky-400">👥</span>
            <h2 className="text-xl font-black tracking-tight text-slate-900 dark:text-white uppercase">Control de Nómina y RRHH</h2>
          </div>
          <p className="text-xs text-slate-500 dark:text-slate-400 font-medium">Gestión de asistencias, pagos, descuentos y directorio de colaboradores.</p>
        </div>
        
        <div className="flex items-center gap-2 flex-wrap justify-center xl:justify-end w-full xl:w-auto">
          {tabInterna === 'liquidacion' && (
            <button onClick={exportarNominaAExcel} className="px-4 py-2.5 bg-emerald-600 hover:bg-emerald-700 text-white font-black text-xs uppercase tracking-wider rounded-xl shadow-md transition-all flex items-center gap-1.5 cursor-pointer">
              <span>📊</span> Exportar
            </button>
          )}

          {tabInterna === 'personal' && (
            <button onClick={() => { limpiarFormTrabajador(); setModalActivo('trabajador'); }} className="px-4 py-2.5 bg-[#117097] hover:bg-[#0a4c68] text-white font-black text-xs uppercase tracking-wider rounded-xl shadow-md transition-all flex items-center gap-1.5 cursor-pointer">
              <span>➕</span> Trabajador
            </button>
          )}

          {tabInterna === 'planilla' && (
            <>
              <button onClick={() => { setFormJornal({ ...formJornal, id_editando: null }); setModalActivo('labor'); }} className="px-4 py-2.5 bg-emerald-700 hover:bg-emerald-800 text-white font-black text-xs uppercase tracking-wider rounded-xl shadow-md transition-all flex items-center gap-1.5 cursor-pointer">
                <span>📝</span> Labor
              </button>
              <button onClick={() => { setFormVale({ ...formVale, id_editando: null }); setModalActivo('vale'); }} className="px-4 py-2.5 bg-rose-600 hover:bg-rose-700 text-white font-black text-xs uppercase tracking-wider rounded-xl shadow-md transition-all flex items-center gap-1.5 cursor-pointer">
                <span>🎟️</span> Vale
              </button>
            </>
          )}
        </div>
      </div>

      {/* 🧭 TABS NAVEGACIÓN */}
      <div className="bg-white dark:bg-slate-800 p-4 rounded-2xl shadow-md border border-slate-200 dark:border-slate-700 flex gap-2 overflow-x-auto w-full transition-colors duration-300">
        <button onClick={() => setTabInterna('planilla')} className={`px-4 py-2 rounded-xl text-[10px] uppercase font-black transition-all whitespace-nowrap cursor-pointer ${tabInterna === 'planilla' ? 'bg-[#117097] text-white shadow' : 'bg-slate-100 dark:bg-slate-700 text-slate-500 dark:text-slate-300 hover:bg-slate-200 dark:hover:bg-slate-600'}`}>📅 Planilla Diaria / Vales</button>
        <button onClick={() => setTabInterna('liquidacion')} className={`px-4 py-2 rounded-xl text-[10px] uppercase font-black transition-all whitespace-nowrap cursor-pointer ${tabInterna === 'liquidacion' ? 'bg-[#117097] text-white shadow' : 'bg-slate-100 dark:bg-slate-700 text-slate-500 dark:text-slate-300 hover:bg-slate-200 dark:hover:bg-slate-600'}`}>💰 Liquidación Central</button>
        <button onClick={() => setTabInterna('personal')} className={`px-4 py-2 rounded-xl text-[10px] uppercase font-black transition-all whitespace-nowrap cursor-pointer ${tabInterna === 'personal' ? 'bg-[#117097] text-white shadow' : 'bg-slate-100 dark:bg-slate-700 text-slate-500 dark:text-slate-300 hover:bg-slate-200 dark:hover:bg-slate-600'}`}>👥 Registro Trabajadores</button>
      </div>

      {/* ------------------------------------------------------------- */}
      {/* VISTA 1: PLANILLA DIARIA */}
      {/* ------------------------------------------------------------- */}
      {tabInterna === 'planilla' && (
        <div className="bg-white dark:bg-slate-800 rounded-3xl shadow-xl overflow-hidden border border-gray-200 dark:border-slate-700 transition-colors duration-300">
          <div className="bg-[#117097] dark:bg-slate-900 py-3 px-5 text-white font-black text-xs uppercase tracking-wider flex justify-between items-center">
            <span>📋 Planilla Acumulada (Labores & Vales)</span>
            <span className="bg-white/20 px-3 py-1 rounded-md text-[10px] font-black">
              Registros: {listaPlanillaUnificada.length}
            </span>
          </div>
          
          <div className="overflow-x-auto max-h-[600px] overflow-y-auto">
            <table className="w-full text-left border-collapse text-[11px]">
              <thead>
                <tr className="bg-slate-100 dark:bg-slate-700/80 text-slate-700 dark:text-slate-300 uppercase font-black border-b dark:border-slate-700 sticky top-0 z-10">
                  <th className="py-2.5 px-3 text-center">Fecha</th>
                  <th className="py-2.5 px-3">Operario / Colaborador</th>
                  <th className="py-2.5 px-3 text-center">Invernadero / Sede</th>
                  <th className="py-2.5 px-3 text-center">Tipo Registro / Labor</th>
                  <th className="py-2.5 px-3">Detalle / Notas</th>
                  <th className="py-2.5 px-3 text-right">Monto (COP)</th>
                  <th className="py-2.5 px-3 text-center">Acciones</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-gray-100 dark:divide-slate-700/60 font-bold text-slate-700 dark:text-slate-300">
                {listaPlanillaUnificada.length === 0 ? (
                  <tr>
                    <td colSpan="7" className="p-6 text-center text-slate-400 dark:text-slate-500 italic font-black">No hay labores ni vales cargados en la planilla esta semana</td>
                  </tr>
                ) : (
                  listaPlanillaUnificada.map((item, idx) => {
                    const esVale = item.tipo_registro === 'VALE';
                    return (
                      <tr key={`${item.tipo_registro}-${item.id}-${idx}`} className={`hover:bg-sky-50 dark:hover:bg-slate-700/50 transition-colors ${idx % 2 === 0 ? 'bg-white dark:bg-slate-800' : 'bg-slate-50/50 dark:bg-slate-800/60'} border-l-4 ${esVale ? 'border-rose-500' : 'border-[#117097]'}`}>
                        <td className="py-2 px-3 text-center font-black text-slate-500 dark:text-slate-400 whitespace-nowrap">{item.fecha}</td>
                        <td className="py-2 px-3 font-black text-slate-900 dark:text-white uppercase text-xs">{item.trabajador_nombre}</td>
                        <td className="py-2 px-3 text-center uppercase font-bold text-[#117097] dark:text-sky-400 text-[10px]">{item.invernadero_nombre}</td>
                        <td className="py-2 px-3 text-center whitespace-nowrap">
                          <span className={`px-2 py-0.5 rounded-md font-black text-[9px] uppercase shadow-sm border ${
                            esVale ? 'bg-rose-50 dark:bg-rose-950/80 text-rose-800 dark:text-rose-400 border-rose-200 dark:border-rose-800' : 'bg-emerald-50 dark:bg-emerald-950/80 text-emerald-800 dark:text-emerald-400 border-emerald-200 dark:border-emerald-800'
                          }`}>
                            {item.concepto_tipo}
                          </span>
                        </td>
                        <td className="py-2 px-3 uppercase text-slate-500 dark:text-slate-400 text-[10px]">{item.detalle_nota}</td>
                        <td className="py-2 px-3 text-right whitespace-nowrap">
                          {esVale ? (
                            <span className="font-black text-rose-600 dark:text-rose-400 text-xs">-{formatoPesos(item.monto_descuento)}</span>
                          ) : (
                            <span className="font-black text-emerald-700 dark:text-emerald-400 text-xs">+{formatoPesos(item.monto_devengado)}</span>
                          )}
                        </td>
                        <td className="py-2 px-3 text-center whitespace-nowrap">
                          <div className="flex gap-1 justify-center">
                            {esVale ? (
                              <>
                                <button onClick={() => prepararEdicionVale(item.objeto_original)} className="p-1 bg-slate-700 dark:bg-slate-600 text-white rounded-lg border border-slate-800 dark:border-slate-500 hover:bg-slate-900 dark:hover:bg-slate-500 transition-all text-[9px] cursor-pointer" title="Editar Vale">✏️</button>
                                <button onClick={() => eliminarValePendiente(item.id)} className="p-1 bg-red-100 dark:bg-red-950/60 text-red-700 dark:text-red-400 rounded-lg border border-red-200 dark:border-red-900 hover:bg-red-600 hover:text-white transition-all text-[9px] cursor-pointer" title="Eliminar Vale">🗑️</button>
                              </>
                            ) : (
                              <>
                                <button onClick={() => prepararEdicionJornal(item.objeto_original)} className="p-1 bg-slate-700 dark:bg-slate-600 text-white rounded-lg border border-slate-800 dark:border-slate-500 hover:bg-slate-900 dark:hover:bg-slate-500 transition-all text-[9px] cursor-pointer" title="Editar Labor">✏️</button>
                                <button onClick={() => eliminarJornalPendiente(item.id)} className="p-1 bg-red-100 dark:bg-red-950/60 text-red-700 dark:text-red-400 rounded-lg border border-red-200 dark:border-red-900 hover:bg-red-600 hover:text-white transition-all text-[9px] cursor-pointer" title="Eliminar Labor">🗑️</button>
                              </>
                            )}
                          </div>
                        </td>
                      </tr>
                    );
                  })
                )}
              </tbody>
            </table>
          </div>
        </div>
      )}

      {/* ------------------------------------------------------------- */}
      {/* VISTA 2: LIQUIDACIÓN CENTRAL (FILAS COMPRIMIDAS) */}
      {/* ------------------------------------------------------------- */}
      {tabInterna === 'liquidacion' && (
        <div className="space-y-6">
          <div className="bg-white dark:bg-slate-800 rounded-3xl shadow-xl overflow-hidden border border-gray-200 dark:border-slate-700 transition-colors duration-300">
            <div className="p-4 bg-[#117097] dark:bg-slate-900 text-white font-black text-xs uppercase tracking-wider flex flex-col md:flex-row md:items-center justify-between gap-4">
              
              <div className="flex gap-2">
                <button 
                  onClick={() => setSubTabLiquidacion('Sábado (Jornalero)')} 
                  className={`px-4 py-2.5 rounded-xl text-[10px] font-black uppercase transition-all flex items-center gap-2 cursor-pointer shadow-sm ${
                    subTabLiquidacion === 'Sábado (Jornalero)' ? 'bg-emerald-600 text-white border-2 border-white ring-2 ring-emerald-400/50 scale-105' : 'bg-slate-800 text-slate-300 hover:bg-slate-700 border border-slate-600'
                  }`}
                >
                  <span>🚜 JORNALEROS (SÁBADO)</span>
                  <span className={`px-2 py-0.5 rounded-full text-[9px] font-black ${subTabLiquidacion === 'Sábado (Jornalero)' ? 'bg-white text-emerald-800' : 'bg-slate-600 text-slate-200'}`}>
                    {listaLiquidacionSabado.length}
                  </span>
                </button>

                <button 
                  onClick={() => setSubTabLiquidacion('Quincenal (Fijo)')} 
                  className={`px-4 py-2.5 rounded-xl text-[10px] font-black uppercase transition-all flex items-center gap-2 cursor-pointer shadow-sm ${
                    subTabLiquidacion === 'Quincenal (Fijo)' ? 'bg-indigo-600 text-white border-2 border-white ring-2 ring-indigo-400/50 scale-105' : 'bg-slate-800 text-slate-300 hover:bg-slate-700 border border-slate-600'
                  }`}
                >
                  <span>📅 FIJOS (QUINCENAL)</span>
                  <span className={`px-2 py-0.5 rounded-full text-[9px] font-black ${subTabLiquidacion === 'Quincenal (Fijo)' ? 'bg-white text-indigo-900' : 'bg-slate-600 text-slate-200'}`}>
                    {listaLiquidacionQuincena.length}
                  </span>
                </button>
              </div>

              <button onClick={resetearCicloCompleto} className="px-4 py-2 bg-slate-800 hover:bg-slate-900 text-slate-200 text-[10px] font-black uppercase rounded-xl transition-colors border border-slate-700 cursor-pointer shadow">🧹 Limpiar Asistencia</button>
            </div>

            <div className="overflow-x-auto max-h-[500px]">
              <table className="w-full text-left border-collapse text-[11px]">
                <thead>
                  <tr className="bg-slate-100 dark:bg-slate-700/80 text-slate-600 dark:text-slate-300 uppercase font-black text-[10px] border-b dark:border-slate-700 sticky top-0 z-10">
                    <th className="py-2.5 px-3">Trabajador / Colaborador</th>
                    <th className="py-2.5 px-3 text-center">Invernadero / Sede</th>
                    <th className="py-2.5 px-3 text-center">Definir Fecha Corte</th>
                    <th className="py-2.5 px-3 text-center w-3/12">Forma de Pago Efectiva / Cuenta</th>
                    <th className="py-2.5 px-3 text-right">Monto Devengado (+)</th>
                    <th className="py-2.5 px-3 text-right">Vales (-)</th>
                    <th className="py-2.5 px-3 text-right">Neto a Pagar (=)</th>
                    <th className="py-2.5 px-3 text-center">Acciones</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-gray-100 dark:divide-slate-700/60 font-bold text-slate-700 dark:text-slate-300">
                  {((subTabLiquidacion === 'Sábado (Jornalero)' && listaLiquidacionSabado.length === 0) || 
                    (subTabLiquidacion === 'Quincenal (Fijo)' && listaLiquidacionQuincena.length === 0)) ? (
                    <tr><td colSpan="8" className="p-6 text-center text-slate-400 dark:text-slate-500 italic font-bold">No hay pagos pendientes de procesar.</td></tr>
                  ) : (
                    (subTabLiquidacion === 'Sábado (Jornalero)' ? listaLiquidacionSabado : listaLiquidacionQuincena).map((item) => (
                      <tr key={item.id} className="bg-white dark:bg-slate-800 hover:bg-sky-50 dark:hover:bg-slate-700/50 transition-colors border-l-4 border-[#117097]">
                        <td className="py-1.5 px-3 font-black uppercase text-slate-900 dark:text-white whitespace-nowrap text-xs">{item.nombre}</td>
                        
                        <td className="py-1 px-3 text-center">
                          <select 
                            className="border border-slate-300 dark:border-slate-600 p-1.5 rounded-lg text-[9px] font-black text-[#117097] dark:text-sky-400 bg-white dark:bg-slate-900 outline-none focus:border-[#117097]"
                            value={invernaderosSeleccionados[item.id] || ''}
                            onChange={e => setInvernaderosSeleccionados({ ...invernaderosSeleccionados, [item.id]: e.target.value })}
                          >
                            <option value="">GENERAL / VARIOS</option>
                            {listaInvernaderos?.filter(inv => inv.activo !== false).map(inv => <option key={inv.id} value={inv.id}>{inv.nombre?.toUpperCase()}</option>)}
                          </select>
                        </td>

                        <td className="py-1 px-3 text-center">
                          <input type="date" className="border border-slate-300 dark:border-slate-600 p-1.5 rounded-lg text-center font-black text-[10px] text-[#117097] dark:text-sky-400 bg-white dark:bg-slate-900 outline-none focus:border-[#117097]" value={fechasCorteQuincena[item.id] || ''} onChange={e => setFechasCorteQuincena({ ...fechasCorteQuincena, [item.id]: e.target.value })} />
                        </td>
                        
                        <td className="py-1 px-3">
                          <div className="flex items-center gap-1.5 min-w-[220px] justify-center mx-auto">
                            <select 
                              className="w-1/2 border border-slate-300 dark:border-slate-600 p-1.5 rounded-lg text-[9px] font-black text-[#117097] dark:text-sky-400 bg-white dark:bg-slate-900 outline-none focus:border-[#117097]" 
                              value={formasPagoModificadas[item.id] || 'Efectivo'} 
                              onChange={e => setFormasPagoModificadas({ ...formasPagoModificadas, [item.id]: e.target.value })}
                            >
                              {formasPagoOpciones.map(fp => <option key={fp} value={fp}>{fp}</option>)}
                            </select>
                            <input 
                              type="text" 
                              className="w-1/2 border border-dashed border-slate-300 dark:border-slate-600 p-1.5 rounded-lg bg-slate-50 dark:bg-slate-900 text-[10px] text-center font-black text-[#117097] dark:text-sky-400 outline-none focus:border-[#117097] placeholder-slate-400" 
                              value={cuentasModificadas[item.id] || ''} 
                              onChange={e => setCuentasModificadas({ ...cuentasModificadas, [item.id]: e.target.value })} 
                              placeholder="N° Cuenta / Celular" 
                            />
                          </div>
                        </td>

                        <td className="py-1.5 px-3 text-right font-black text-xs text-emerald-700 dark:text-emerald-400 whitespace-nowrap">+{formatoPesos(item.totalGanado)}</td>
                        <td className="py-1.5 px-3 text-right font-black text-xs text-rose-600 dark:text-rose-400 whitespace-nowrap">-{formatoPesos(item.totalVales)}</td>
                        <td className="py-1.5 px-3 text-right font-black text-xs text-[#117097] dark:text-sky-400 whitespace-nowrap">{formatoPesos(item.netoPagar)}</td>
                        
                        <td className="py-1.5 px-3">
                          <div className="flex justify-center items-center gap-1">
                            <button type="button" onClick={() => setTrabajadorDetalleEdicion(item)} className="p-1 bg-slate-700 dark:bg-slate-600 text-white rounded-lg hover:bg-slate-900 dark:hover:bg-slate-500 transition-all text-[9px] font-black cursor-pointer shadow-sm" title="Ver Registros">🔍</button>
                            <button type="button" onClick={() => generarComprobanteNominaPDF(item)} className="p-1 bg-[#0a4c68] text-white rounded-lg hover:bg-black transition-all text-[9px] font-black cursor-pointer shadow-sm" title="Imprimir PDF">🖨️</button>
                            <button type="button" onClick={() => pagarNominaTrabajador(item)} className="px-2.5 py-1 bg-emerald-600 text-white rounded-lg hover:bg-emerald-700 transition-all text-[9px] font-black uppercase cursor-pointer shadow-sm">💵 Pagar</button>
                            <button type="button" onClick={() => descartarConsolidadoTrabajador(item.id, item.nombre)} className="p-1 bg-red-100 dark:bg-red-950/60 text-red-700 dark:text-red-400 rounded-lg hover:bg-red-600 hover:text-white transition-all text-[9px] border border-red-200 dark:border-red-900 cursor-pointer" title="Descartar">🗑️</button>
                          </div>
                        </td>
                      </tr>
                    ))
                  )}
                </tbody>
              </table>
            </div>
          </div>

          <div className="bg-white dark:bg-slate-800 rounded-3xl shadow-xl overflow-hidden border border-gray-200 dark:border-slate-700 transition-colors duration-300">
            <div className="p-4 bg-slate-800 dark:bg-slate-900 text-white font-black text-[11px] uppercase tracking-wider flex justify-between items-center flex-wrap gap-2">
              <span>📚 Historial de Periodos Liquidados Anteriormente</span>
            </div>
            <div className="overflow-x-auto max-h-[500px] overflow-y-auto">
              <table className="w-full text-left border-collapse text-[11px]">
                <thead>
                  <tr className="bg-slate-100 dark:bg-slate-700/80 text-slate-600 dark:text-slate-300 uppercase font-black border-b dark:border-slate-700 sticky top-0 z-10 text-[10px]">
                    <th className="py-2.5 px-3 text-center">Comp. N°</th>
                    <th className="py-2.5 px-3">Trabajador / Colaborador</th>
                    <th className="py-2.5 px-3 text-center">Invernadero</th>
                    <th className="py-2.5 px-3 text-center">Fecha Periodo</th>
                    <th className="py-2.5 px-3 text-center">Forma Pago Efectiva</th>
                    <th className="py-2.5 px-3 text-right">Devengado (+)</th>
                    <th className="py-2.5 px-3 text-right">Vales (-)</th>
                    <th className="py-2.5 px-3 text-right">Neto Entregado (=)</th>
                    <th className="py-2.5 px-3 text-center">Acciones</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-gray-100 dark:divide-slate-700/60 font-bold text-slate-700 dark:text-slate-300">
                  {pagosHistoricos.filter(p => (trabajadores.find(t => t.id === p.trabajador_id)?.tipo_pago || 'Sábado (Jornalero)') === subTabLiquidacion).map(p => {
                    const t = trabajadores.find(trab => trab.id === p.trabajador_id);
                    return (
                      <tr key={p.id} className="hover:bg-slate-50 dark:hover:bg-slate-700/50 transition-colors border-l-4 border-slate-300 dark:border-slate-600">
                        <td className="py-1.5 px-3 text-center font-black text-slate-500 dark:text-slate-400 whitespace-nowrap">NOM-{String(p.id).padStart(4, '0')}</td>
                        <td className="py-1.5 px-3 font-black text-slate-900 dark:text-white uppercase whitespace-nowrap">{t?.nombre_completo}</td>
                        <td className="py-1.5 px-3 text-center font-black text-[#117097] dark:text-sky-400 uppercase text-[10px]">{p.invernadero_nombre || 'GENERAL / VARIOS'}</td>
                        <td className="py-1.5 px-3 text-center text-slate-600 dark:text-slate-400 whitespace-nowrap">{p.fecha_pago}</td>
                        <td className="py-1.5 px-3 text-center text-slate-500 dark:text-slate-400 font-bold uppercase text-[10px]">
                          {p.forma_pago_efectiva || 'Efectivo'} {p.numero_cuenta_efectivo ? `(#${p.numero_cuenta_efectivo})` : ''}
                        </td>
                        <td className="py-1.5 px-3 text-right text-emerald-700 dark:text-emerald-400 whitespace-nowrap">+{formatoPesos(p.monto_devengado)}</td>
                        <td className="py-1.5 px-3 text-rose-500 dark:text-rose-400 text-right whitespace-nowrap">-{formatoPesos(p.monto_vales)}</td>
                        <td className="py-1.5 px-3 text-right text-slate-900 dark:text-white font-black text-xs whitespace-nowrap">{formatoPesos(p.monto_pagado)}</td>
                        <td className="py-1.5 px-3 text-center whitespace-nowrap">
                          <div className="flex justify-center gap-1">
                            <button 
                              type="button" 
                              onClick={() => generarComprobanteNominaPDF({ id: p.trabajador_id, nombre: t?.nombre_completo, cedula: t?.cedula, tipo_pago: subTabLiquidacion, netoPagar: p.monto_pagado, totalGanado: p.monto_devengado, totalVales: p.monto_vales, dias_respaldo: p.dias_liquidados, invernadero_nombre: p.invernadero_nombre, jornalesDetalle: p.jornalesDetalle, valesDetalle: p.valesDetalle }, p.fecha_pago, p.forma_pago_efectiva, p.numero_cuenta_efectivo, p.id)} 
                              className="px-2 py-1 bg-slate-800 dark:bg-slate-700 text-white rounded-lg hover:bg-black dark:hover:bg-slate-600 text-[9px] font-black cursor-pointer shadow-sm" title="Reimprimir Voucher"
                            >🖨️ PDF</button>
                            <button type="button" onClick={() => setPagoEditando(p)} className="p-1 bg-slate-200 dark:bg-slate-700 text-slate-700 dark:text-slate-200 rounded-lg border dark:border-slate-600 hover:bg-slate-300 dark:hover:bg-slate-600 text-[9px] cursor-pointer" title="Editar Pago">✏️</button>
                            <button type="button" onClick={() => eliminarPagoHistorico(p)} className="p-1 bg-red-100 dark:bg-red-950/60 text-red-700 dark:text-red-400 rounded-lg border hover:bg-red-600 hover:text-white border-red-200 dark:border-red-900 text-[9px] cursor-pointer" title="Anular Pago">🗑️</button>
                          </div>
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

      {/* ------------------------------------------------------------- */}
      {/* VISTA 3: REGISTRO DE TRABAJADORES */}
      {/* ------------------------------------------------------------- */}
      {tabInterna === 'personal' && (
        <div className="bg-white dark:bg-slate-800 rounded-3xl shadow-xl overflow-hidden border border-gray-200 dark:border-slate-700 transition-colors duration-300">
          <div className="p-4 bg-[#117097] dark:bg-slate-900 text-white font-black text-xs flex justify-between items-center uppercase tracking-wider">
            <span>👥 Directorio de Colaboradores ({trabajadores.length})</span>
          </div>
          <div className="overflow-x-auto max-h-[600px] overflow-y-auto">
            <table className="w-full text-left border-collapse text-[11px]">
              <thead>
                <tr className="bg-slate-100 dark:bg-slate-700/80 text-slate-700 dark:text-slate-300 uppercase font-black text-[10px] border-b dark:border-slate-700 sticky top-0">
                  <th className="py-2.5 px-3">Operario / Cédula</th>
                  <th className="py-2.5 px-3">Contacto</th>
                  <th className="py-2.5 px-3">Forma de Pago / Cuenta</th>
                  <th className="py-2.5 px-3 text-center">F. Ingreso</th>
                  <th className="py-2.5 px-3 text-right">Frecuencia / Base</th>
                  <th className="py-2.5 px-3 text-center">Acciones</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-gray-100 dark:divide-slate-700/60 font-bold text-slate-700 dark:text-slate-300">
                {trabajadores.length === 0 ? (
                  <tr><td colSpan="6" className="p-6 text-center text-slate-400 italic">No hay operarios registrados</td></tr>
                ) : (
                  trabajadores.map((t, idx) => (
                    <tr key={t.id} className={`${idx % 2 === 0 ? 'bg-white dark:bg-slate-800' : 'bg-slate-50/50 dark:bg-slate-800/60'} hover:bg-sky-50 dark:hover:bg-slate-700/50 transition-colors border-l-4 border-[#117097]`}>
                      <td className="py-2 px-3 uppercase">
                        <p className="font-black text-slate-900 dark:text-white text-xs">{t.nombre_completo}</p>
                        <p className="text-[10px] text-gray-500 dark:text-slate-400 font-bold">C.C. {t.cedula || 'N/R'}</p>
                      </td>
                      <td className="py-2 px-3">
                        <p className="text-slate-800 dark:text-slate-200 font-black text-[10px]">📞 {t.telefono || 'N/R'}</p>
                        {t.email && <p className="text-[9px] text-[#117097] dark:text-sky-400 font-bold lowercase">✉️ {t.email}</p>}
                      </td>
                      <td className="py-2 px-3">
                        <span className="inline-block bg-slate-100 dark:bg-slate-700 border border-slate-200 dark:border-slate-600 text-slate-700 dark:text-slate-300 text-[9px] px-2 py-0.5 rounded-md font-black uppercase shadow-sm">
                          💳 {t.forma_pago_predeterminada || 'Efectivo'}
                        </span>
                        {t.numero_cuenta_predeterminado && (
                          <p className="text-[9px] text-[#117097] dark:text-sky-400 font-black mt-0.5">#{t.numero_cuenta_predeterminado}</p>
                        )}
                      </td>
                      <td className="py-2 px-3 text-center text-slate-500 dark:text-slate-400 font-black text-[10px]">
                        {t.fecha_registro || '---'}
                      </td>
                      <td className="py-2 px-3 text-right">
                        <span className="block text-[8px] uppercase font-black text-slate-400">
                          {t.tipo_pago === 'Quincenal (Fijo)' ? 'Sueldo Fijo' : 'Jornal Base'}
                        </span>
                        <span className="font-black text-[#117097] dark:text-sky-400 text-xs inline-block">
                          {formatoPesos(t.pago_jornal_base)}
                        </span>
                      </td>
                      <td className="py-2 px-3 text-center">
                        <div className="flex gap-1 justify-center">
                          <button onClick={() => prepararEdicionTrabajador(t)} className="p-1 bg-slate-700 dark:bg-slate-600 text-white rounded-lg hover:bg-slate-900 dark:hover:bg-slate-500 transition-all text-[9px] font-black shadow-sm cursor-pointer" title="Editar">✏️</button>
                          <button onClick={() => eliminarTrabajadorLogico(t.id, t.nombre_completo)} className="p-1 bg-red-100 dark:bg-red-950/60 text-red-700 dark:text-red-400 rounded-lg hover:bg-red-600 hover:text-white transition-all text-[9px] border border-red-200 dark:border-red-900 cursor-pointer" title="Inactivar">🗑️</button>
                        </div>
                      </td>
                    </tr>
                  ))
                )}
              </tbody>
            </table>
          </div>
        </div>
      )}

      {/* MODALES FLOTANTES */}
      {modalActivo === 'trabajador' && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 backdrop-blur-sm p-4 overflow-y-auto">
          <div className="bg-white dark:bg-slate-800 border border-slate-200 dark:border-slate-700 w-full max-w-lg rounded-3xl shadow-2xl overflow-hidden animate-in fade-in zoom-in-95 my-8 text-slate-800 dark:text-slate-200">
            <div className="p-5 bg-slate-900 text-white border-b border-slate-800 flex justify-between items-center">
              <h3 className="font-black uppercase tracking-tight flex items-center gap-2"><span>👥</span> {formTrabajador.id_editando ? 'Editar Trabajador' : 'Nuevo Trabajador'}</h3>
              <button onClick={() => setModalActivo(null)} className="text-slate-400 hover:text-white font-black px-3 py-1 bg-slate-800 rounded-xl cursor-pointer">✕</button>
            </div>
            <form onSubmit={registrarTrabajador} className="p-6 space-y-4">
              <div>
                <label className="text-[10px] font-black text-slate-500 dark:text-slate-400 uppercase tracking-wider">Nombre Completo *</label>
                <input type="text" className="w-full bg-slate-50 dark:bg-slate-900 border border-slate-200 dark:border-slate-700 text-slate-800 dark:text-white p-2.5 rounded-xl font-bold text-xs uppercase outline-none focus:border-[#117097] mt-1 placeholder-slate-400" value={formTrabajador.nombre_completo} onChange={e => setFormTrabajador({...formTrabajador, nombre_completo: e.target.value})} required />
              </div>
              <div className="grid grid-cols-2 gap-4">
                <div>
                  <label className="text-[10px] font-black text-slate-500 dark:text-slate-400 uppercase tracking-wider">Cédula *</label>
                  <input type="text" className="w-full bg-slate-50 dark:bg-slate-900 border border-slate-200 dark:border-slate-700 text-slate-800 dark:text-white p-2.5 rounded-xl font-bold text-xs outline-none focus:border-[#117097] mt-1 placeholder-slate-400" value={formTrabajador.cedula} onChange={e => setFormTrabajador({...formTrabajador, cedula: e.target.value})} required placeholder="Sin puntos" />
                </div>
                <div>
                  <label className="text-[10px] font-black text-slate-500 dark:text-slate-400 uppercase tracking-wider">Teléfono</label>
                  <input type="text" className="w-full bg-slate-50 dark:bg-slate-900 border border-slate-200 dark:border-slate-700 text-slate-800 dark:text-white p-2.5 rounded-xl font-bold text-xs outline-none focus:border-[#117097] mt-1 placeholder-slate-400" value={formTrabajador.telefono} onChange={e => setFormTrabajador({...formTrabajador, telefono: e.target.value})} />
                </div>
              </div>
              <div>
                <label className="text-[10px] font-black text-slate-500 dark:text-slate-400 uppercase tracking-wider">Correo Electrónico</label>
                <input type="email" className="w-full bg-slate-50 dark:bg-slate-900 border border-slate-200 dark:border-slate-700 text-slate-800 dark:text-white p-2.5 rounded-xl font-bold text-xs lowercase outline-none focus:border-[#117097] mt-1 placeholder-slate-400" value={formTrabajador.email || ''} onChange={e => setFormTrabajador({...formTrabajador, email: e.target.value})} />
              </div>
              <div className="grid grid-cols-2 gap-4">
                <div>
                  <label className="text-[10px] font-black text-slate-500 dark:text-slate-400 uppercase tracking-wider">Frecuencia Pago *</label>
                  <select className="w-full bg-slate-50 dark:bg-slate-900 border border-slate-200 dark:border-slate-700 text-slate-800 dark:text-white p-2.5 rounded-xl font-bold text-xs outline-none focus:border-[#117097] mt-1" value={formTrabajador.tipo_pago} onChange={e => setFormTrabajador({...formTrabajador, tipo_pago: e.target.value})}>
                    {tiposPago.map(p => <option key={p} value={p}>{p}</option>)}
                  </select>
                </div>
                <div>
                  <label className="text-[10px] font-black text-slate-500 dark:text-slate-400 uppercase tracking-wider">Fecha Ingreso *</label>
                  <input type="date" className="w-full bg-slate-50 dark:bg-slate-900 border border-slate-200 dark:border-slate-700 text-slate-800 dark:text-white p-2.5 rounded-xl font-bold text-xs outline-none focus:border-[#117097] mt-1" value={formTrabajador.fecha_registro} onChange={e => setFormTrabajador({...formTrabajador, fecha_registro: e.target.value})} required />
                </div>
              </div>
              <div>
                <label className="text-[10px] font-black text-[#117097] dark:text-sky-400 uppercase tracking-wider">
                  {formTrabajador.tipo_pago === 'Quincenal (Fijo)' ? 'Sueldo Quincenal Fijo *' : 'Tarifa Jornal Diario Base *'}
                </label>
                <input type="text" className="w-full border-2 border-[#117097]/50 dark:border-sky-700 bg-white dark:bg-slate-900 text-slate-800 dark:text-white p-2.5 rounded-xl font-black text-sm text-[#117097] dark:text-sky-400 outline-none focus:border-[#117097] mt-1 placeholder-slate-400" value={formTrabajador.pago_jornal_base ? formatoPesos(formTrabajador.pago_jornal_base) : ''} onChange={e => setFormTrabajador({...formTrabajador, pago_jornal_base: e.target.value.replace(/\D/g, '')})} required placeholder="$ 0" />
              </div>
              <div className="bg-slate-50 dark:bg-slate-900/60 p-4 rounded-xl border border-slate-200 dark:border-slate-700">
                <p className="text-[10px] font-black text-slate-500 dark:text-slate-400 uppercase tracking-wider mb-2">Información de Transferencia</p>
                <div className="grid grid-cols-2 gap-4">
                  <div>
                    <label className="text-[9px] font-black text-slate-400 uppercase">Banco</label>
                    <select className="w-full border border-slate-200 dark:border-slate-600 bg-white dark:bg-slate-900 text-slate-800 dark:text-white p-2 rounded-lg text-xs font-bold mt-1" value={formTrabajador.forma_pago_predeterminada} onChange={e => setFormTrabajador({...formTrabajador, forma_pago_predeterminada: e.target.value})}>
                      {formasPagoOpciones.map(fp => <option key={fp} value={fp}>{fp}</option>)}
                    </select>
                  </div>
                  <div>
                    <label className="text-[9px] font-black text-slate-400 uppercase">N° Cuenta</label>
                    <input type="text" className="w-full border border-slate-200 dark:border-slate-600 bg-white dark:bg-slate-900 text-slate-800 dark:text-white p-2 rounded-lg text-xs font-black mt-1 placeholder-slate-400" value={formTrabajador.numero_cuenta_predeterminado} onChange={e => setFormTrabajador({...formTrabajador, numero_cuenta_predeterminado: e.target.value})} />
                  </div>
                </div>
              </div>
              <div className="flex justify-end gap-3 pt-4 border-t border-slate-100 dark:border-slate-700">
                <button type="button" onClick={() => setModalActivo(null)} className="px-5 py-2.5 bg-slate-200 dark:bg-slate-700 hover:bg-slate-300 dark:hover:bg-slate-600 text-slate-700 dark:text-slate-200 font-black text-xs uppercase rounded-xl">Cancelar</button>
                <button type="submit" className="px-6 py-2.5 bg-[#117097] text-white font-black text-xs uppercase rounded-xl shadow-lg">💾 Guardar</button>
              </div>
            </form>
          </div>
        </div>
      )}

      {modalActivo === 'labor' && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 backdrop-blur-sm p-4 overflow-y-auto">
          <div className="bg-white dark:bg-slate-800 border border-slate-200 dark:border-slate-700 w-full max-w-lg rounded-3xl shadow-2xl overflow-hidden animate-in fade-in zoom-in-95 my-8 text-slate-800 dark:text-slate-200">
            <div className="p-5 bg-emerald-800 text-white border-b border-emerald-900 flex justify-between items-center">
              <h3 className="font-black uppercase tracking-tight flex items-center gap-2"><span>📝</span> {formJornal.id_editando ? 'Editar Labor' : 'Registrar Asistencia / Labor'}</h3>
              <button onClick={() => setModalActivo(null)} className="text-emerald-300 hover:text-white font-black px-3 py-1 bg-emerald-900 rounded-xl cursor-pointer">✕</button>
            </div>
            <form onSubmit={registrarJornalDiario} className="p-6 space-y-4">
              <div className="grid grid-cols-2 gap-4">
                <div>
                  <label className="text-[10px] font-black text-slate-500 dark:text-slate-400 uppercase">Fecha *</label>
                  <input type="date" className="w-full bg-slate-50 dark:bg-slate-900 border border-slate-200 dark:border-slate-700 text-slate-800 dark:text-white p-2.5 rounded-xl font-bold text-xs mt-1" value={formJornal.fecha_labor} onChange={e => setFormJornal({...formJornal, fecha_labor: e.target.value})} required />
                </div>
                <div>
                  <label className="text-[10px] font-black text-slate-500 dark:text-slate-400 uppercase">Invernadero</label>
                  <select className="w-full bg-slate-50 dark:bg-slate-900 border border-slate-200 dark:border-slate-700 text-slate-800 dark:text-white p-2.5 rounded-xl font-bold text-xs uppercase mt-1" value={formJornal.invernadero_id} onChange={e => setFormJornal({...formJornal, invernadero_id: e.target.value})}>
                    <option value="">GENERAL / VARIOS</option>
                    {listaInvernaderos?.filter(inv => inv.activo !== false).map(inv => <option key={inv.id} value={inv.id}>{inv.nombre?.toUpperCase()}</option>)}
                  </select>
                </div>
              </div>
              <div>
                <label className="text-[10px] font-black text-slate-500 dark:text-slate-400 uppercase">Trabajador *</label>
                <select className="w-full bg-slate-50 dark:bg-slate-900 border border-slate-200 dark:border-slate-700 text-slate-800 dark:text-white p-2.5 rounded-xl font-bold text-xs uppercase mt-1" value={formJornal.trabajador_id} onChange={e => setFormJornal({...formJornal, trabajador_id: e.target.value, tipo_labor: '', valor_pagar: ''})} required>
                  <option value="">Seleccione...</option>
                  {trabajadores.map(t => <option key={t.id} value={t.id}>{t.nombre_completo}</option>)}
                </select>
              </div>
              <div className="grid grid-cols-2 gap-4">
                <div>
                  <label className="text-[10px] font-black text-slate-500 dark:text-slate-400 uppercase">Tipo Labor *</label>
                  <select className="w-full bg-slate-50 dark:bg-slate-900 border border-slate-200 dark:border-slate-700 text-slate-800 dark:text-white p-2.5 rounded-xl font-black text-xs uppercase text-emerald-800 dark:text-emerald-400 mt-1 focus:border-emerald-600" value={formJornal.tipo_labor} onChange={e => {
                    const nTipo = e.target.value;
                    const base = trabajadores.find(t => t.id === formJornal.trabajador_id)?.pago_jornal_base || '';
                    setFormJornal({...formJornal, tipo_labor: nTipo, valor_pagar: nTipo === 'JORNAL' ? base : ''});
                  }} required>
                    <option value="">-- SELECCIONE --</option>
                    {(!trabajadores.find(t=>t.id===formJornal.trabajador_id) || trabajadores.find(t=>t.id===formJornal.trabajador_id).tipo_pago !== 'Quincenal (Fijo)') && <option value="JORNAL">JORNAL COMPLETO</option>}
                    <option value="HORAS">⏱️ POR HORAS</option>
                    <option value="CONTRATO">CONTRATO / TAREA</option>
                  </select>
                </div>
                {formJornal.tipo_labor !== 'HORAS' && (
                  <div>
                    <label className="text-[10px] font-black text-emerald-700 dark:text-emerald-400 uppercase">Total (COP)</label>
                    <input type="text" className="w-full bg-white dark:bg-slate-900 border-2 border-emerald-300 dark:border-emerald-700 p-2.5 rounded-xl font-black text-sm text-emerald-800 dark:text-emerald-300 mt-1 placeholder-slate-400" value={formJornal.valor_pagar ? formatoPesos(formJornal.valor_pagar) : ''} onChange={e => setFormJornal({...formJornal, valor_pagar: e.target.value.replace(/\D/g, '')})} placeholder="$ 0" />
                  </div>
                )}
              </div>

              {formJornal.tipo_labor === 'HORAS' && (
                <div className="bg-amber-50 dark:bg-amber-950/40 p-4 rounded-xl border border-amber-200 dark:border-amber-800 grid grid-cols-3 gap-3">
                  <div>
                    <label className="text-[9px] font-black text-amber-900 dark:text-amber-400 uppercase">Tarifa Hora</label>
                    <input type="text" className="w-full border border-amber-300 dark:border-amber-700 bg-white dark:bg-slate-900 text-slate-800 dark:text-white p-2 rounded-lg text-xs font-black mt-1 placeholder-slate-400" value={formJornal.tarifa_hora ? formatoPesos(formJornal.tarifa_hora) : ''} onChange={e => {
                      const t = e.target.value.replace(/\D/g, '');
                      setFormJornal({...formJornal, tarifa_hora: t, valor_pagar: (parseFloat(t||0) * parseFloat(formJornal.cantidad_horas||0)) || ''});
                    }} />
                  </div>
                  <div>
                    <label className="text-[9px] font-black text-amber-900 dark:text-amber-400 uppercase">Horas</label>
                    <input type="number" step="0.5" className="w-full border border-amber-300 dark:border-amber-700 bg-white dark:bg-slate-900 text-slate-800 dark:text-white p-2 rounded-lg text-xs font-black mt-1" value={formJornal.cantidad_horas} onChange={e => {
                      const c = parseFloat(e.target.value)||0;
                      setFormJornal({...formJornal, cantidad_horas: e.target.value, valor_pagar: (parseFloat(formJornal.tarifa_hora||0) * c) || ''});
                    }} />
                  </div>
                  <div>
                    <label className="text-[9px] font-black text-amber-900 dark:text-amber-400 uppercase">Total</label>
                    <div className="w-full border border-amber-300 dark:border-amber-700 bg-white dark:bg-slate-900 p-2 rounded-lg text-xs font-black mt-1 text-center text-emerald-700 dark:text-emerald-400">{formatoPesos(formJornal.valor_pagar)}</div>
                  </div>
                </div>
              )}

              <div>
                <label className="text-[10px] font-black text-slate-500 dark:text-slate-400 uppercase">Notas / Detalle</label>
                <input type="text" className="w-full bg-slate-50 dark:bg-slate-900 border border-slate-200 dark:border-slate-700 text-slate-800 dark:text-white p-2.5 rounded-xl font-bold text-xs uppercase mt-1 placeholder-slate-400" value={formJornal.observaciones} onChange={e => setFormJornal({...formJornal, observaciones: e.target.value})} placeholder="Ej: Amarre" />
              </div>
              <div className="flex justify-end gap-3 pt-4 border-t border-slate-100 dark:border-slate-700">
                <button type="button" onClick={() => setModalActivo(null)} className="px-5 py-2.5 bg-slate-200 dark:bg-slate-700 hover:bg-slate-300 dark:hover:bg-slate-600 text-slate-700 dark:text-slate-200 font-black text-xs uppercase rounded-xl">Cancelar</button>
                <button type="submit" className="px-6 py-2.5 bg-emerald-700 text-white font-black text-xs uppercase rounded-xl shadow-lg">✓ Guardar</button>
              </div>
            </form>
          </div>
        </div>
      )}

      {modalActivo === 'vale' && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 backdrop-blur-sm p-4 overflow-y-auto">
          <div className="bg-white dark:bg-slate-800 border border-slate-200 dark:border-slate-700 w-full max-w-lg rounded-3xl shadow-2xl overflow-hidden animate-in fade-in zoom-in-95 my-8 text-slate-800 dark:text-slate-200">
            <div className="p-5 bg-rose-800 text-white border-b border-rose-900 flex justify-between items-center">
              <h3 className="font-black uppercase tracking-tight flex items-center gap-2"><span>🎟️</span> {formVale.id_editando ? 'Editar Vale' : 'Registrar Adelanto'}</h3>
              <button onClick={() => setModalActivo(null)} className="text-rose-300 hover:text-white font-black px-3 py-1 bg-rose-900 rounded-xl cursor-pointer">✕</button>
            </div>
            <form onSubmit={registrarValeAdelanto} className="p-6 space-y-4">
              <div className="grid grid-cols-2 gap-4">
                <div>
                  <label className="text-[10px] font-black text-slate-500 dark:text-slate-400 uppercase">Fecha Vale *</label>
                  <input type="date" className="w-full bg-slate-50 dark:bg-slate-900 border border-slate-200 dark:border-slate-700 text-slate-800 dark:text-white p-2.5 rounded-xl font-bold text-xs mt-1" value={formVale.fecha_vale} onChange={e => setFormVale({...formVale, fecha_vale: e.target.value})} required />
                </div>
                <div>
                  <label className="text-[10px] font-black text-rose-700 dark:text-rose-400 uppercase">Monto Descuento *</label>
                  <input type="text" className="w-full bg-white dark:bg-slate-900 border-2 border-rose-300 dark:border-rose-700 p-2.5 rounded-xl font-black text-sm text-rose-800 dark:text-rose-300 mt-1 placeholder-slate-400" value={formVale.monto_vale ? formatoPesos(formVale.monto_vale) : ''} onChange={e => setFormVale({...formVale, monto_vale: e.target.value.replace(/\D/g, '')})} placeholder="$ 0" required />
                </div>
              </div>
              <div>
                <label className="text-[10px] font-black text-slate-500 dark:text-slate-400 uppercase">Trabajador *</label>
                <select className="w-full bg-slate-50 dark:bg-slate-900 border border-slate-200 dark:border-slate-700 text-slate-800 dark:text-white p-2.5 rounded-xl font-bold text-xs uppercase mt-1" value={formVale.trabajador_id} onChange={e => setFormVale({...formVale, trabajador_id: e.target.value})} required>
                  <option value="">Seleccione...</option>
                  {trabajadores.map(t => <option key={t.id} value={t.id}>{t.nombre_completo}</option>)}
                </select>
              </div>
              <div>
                <label className="text-[10px] font-black text-slate-500 dark:text-slate-400 uppercase">Motivo</label>
                <input type="text" className="w-full bg-slate-50 dark:bg-slate-900 border border-slate-200 dark:border-slate-700 text-slate-800 dark:text-white p-2.5 rounded-xl font-bold text-xs uppercase mt-1 placeholder-slate-400" value={formVale.motivo_nota} onChange={e => setFormVale({...formVale, motivo_nota: e.target.value})} placeholder="Ej: Préstamo personal" required />
              </div>
              <div className="flex justify-end gap-3 pt-4 border-t border-slate-100 dark:border-slate-700">
                <button type="button" onClick={() => setModalActivo(null)} className="px-5 py-2.5 bg-slate-200 dark:bg-slate-700 hover:bg-slate-300 dark:hover:bg-slate-600 text-slate-700 dark:text-slate-200 font-black text-xs uppercase rounded-xl">Cancelar</button>
                <button type="submit" className="px-6 py-2.5 bg-rose-700 text-white font-black text-xs uppercase rounded-xl shadow-lg">💸 Aplicar</button>
              </div>
            </form>
          </div>
        </div>
      )}

      {pagoEditando && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 backdrop-blur-sm p-4 overflow-y-auto">
          <div className="bg-white dark:bg-slate-800 border border-slate-200 dark:border-slate-700 w-full max-w-lg rounded-3xl shadow-2xl overflow-hidden animate-in fade-in zoom-in-95 my-8 text-slate-800 dark:text-slate-200">
            <div className="p-5 bg-amber-600 text-white flex justify-between items-center">
              <h3 className="font-black uppercase tracking-tight flex items-center gap-2"><span>✏️</span> Modificar Pago NOM-{String(pagoEditando.id).padStart(4, '0')}</h3>
              <button onClick={() => setPagoEditando(null)} className="text-amber-200 hover:text-white font-black px-3 py-1 bg-amber-800 rounded-xl cursor-pointer">✕</button>
            </div>
            <div className="p-6 space-y-4">
              <div className="grid grid-cols-2 gap-4">
                <div><label className="text-[10px] font-black text-slate-500 dark:text-slate-400 uppercase">Fecha Pago</label><input type="date" className="w-full bg-slate-50 dark:bg-slate-900 border border-slate-200 dark:border-slate-700 text-slate-800 dark:text-white p-2.5 rounded-xl font-bold text-xs mt-1" value={pagoEditando.fecha_pago} onChange={e => setPagoEditando({ ...pagoEditando, fecha_pago: e.target.value })} /></div>
                <div><label className="text-[10px] font-black text-slate-500 dark:text-slate-400 uppercase">Invernadero</label><input type="text" className="w-full bg-slate-50 dark:bg-slate-900 border border-slate-200 dark:border-slate-700 text-slate-800 dark:text-white p-2.5 rounded-xl font-bold text-xs uppercase mt-1" value={pagoEditando.invernadero_nombre} onChange={e => setPagoEditando({ ...pagoEditando, invernadero_nombre: e.target.value.toUpperCase() })} /></div>
              </div>
              <div className="grid grid-cols-2 gap-4">
                <div><label className="text-[10px] font-black text-slate-500 dark:text-slate-400 uppercase">Forma Pago</label><select className="w-full bg-slate-50 dark:bg-slate-900 border border-slate-200 dark:border-slate-700 text-slate-800 dark:text-white p-2.5 rounded-xl font-bold text-xs mt-1" value={pagoEditando.forma_pago_efectiva} onChange={e => setPagoEditando({ ...pagoEditando, forma_pago_efectiva: e.target.value })}>{formasPagoOpciones.map(fp => <option key={fp} value={fp}>{fp}</option>)}</select></div>
                <div><label className="text-[10px] font-black text-slate-500 dark:text-slate-400 uppercase">N° Cuenta</label><input type="text" className="w-full bg-slate-50 dark:bg-slate-900 border border-slate-200 dark:border-slate-700 text-slate-800 dark:text-white p-2.5 rounded-xl font-bold text-xs mt-1 placeholder-slate-400" value={pagoEditando.numero_cuenta_efectivo || ''} onChange={e => setPagoEditando({ ...pagoEditando, numero_cuenta_efectivo: e.target.value })} /></div>
              </div>
              <div className="grid grid-cols-2 gap-4">
                <div><label className="text-[10px] font-black text-emerald-700 dark:text-emerald-400 uppercase">Devengado (+)</label><input type="text" className="w-full bg-white dark:bg-slate-900 border-2 border-emerald-300 dark:border-emerald-700 p-2.5 rounded-xl font-black text-emerald-800 dark:text-emerald-300 text-sm mt-1" value={pagoEditando.monto_devengado ? formatoPesos(pagoEditando.monto_devengado) : ''} onChange={e => setPagoEditando({ ...pagoEditando, monto_devengado: e.target.value.replace(/\D/g, '') })} /></div>
                <div><label className="text-[10px] font-black text-rose-700 dark:text-rose-400 uppercase">Vales (-)</label><input type="text" className="w-full bg-white dark:bg-slate-900 border-2 border-rose-300 dark:border-rose-700 p-2.5 rounded-xl font-black text-rose-800 dark:text-rose-300 text-sm mt-1" value={pagoEditando.monto_vales ? formatoPesos(pagoEditando.monto_vales) : ''} onChange={e => setPagoEditando({ ...pagoEditando, monto_vales: e.target.value.replace(/\D/g, '') })} /></div>
              </div>
              <div className="flex justify-end gap-3 pt-4 border-t border-slate-100 dark:border-slate-700">
                <button onClick={() => setPagoEditando(null)} className="px-5 py-2.5 bg-slate-200 dark:bg-slate-700 hover:bg-slate-300 dark:hover:bg-slate-600 text-slate-700 dark:text-slate-200 font-black text-xs uppercase rounded-xl">Cancelar</button>
                <button onClick={guardarEdicionPagoHistorico} className="px-6 py-2.5 bg-emerald-700 text-white font-black text-xs uppercase rounded-xl shadow-lg">💾 Guardar</button>
              </div>
            </div>
          </div>
        </div>
      )}

      {trabajadorDetalleEdicion && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 backdrop-blur-sm p-4 overflow-y-auto">
          <div className="bg-white dark:bg-slate-800 border border-slate-200 dark:border-slate-700 w-full max-w-3xl rounded-3xl shadow-2xl overflow-hidden animate-in fade-in zoom-in-95 my-8 text-slate-800 dark:text-slate-200">
            <div className="p-5 bg-sky-800 text-white flex justify-between items-center">
              <h3 className="font-black uppercase tracking-tight flex items-center gap-2"><span>🔍</span> Desglose: {trabajadorDetalleEdicion.nombre}</h3>
              <button onClick={() => setTrabajadorDetalleEdicion(null)} className="text-sky-200 hover:text-white font-black px-3 py-1 bg-sky-900 rounded-xl cursor-pointer">✕</button>
            </div>
            <div className="p-6">
              <div className="max-h-80 overflow-y-auto border border-slate-200 dark:border-slate-700 rounded-xl">
                <table className="w-full text-left text-xs">
                  <thead className="bg-slate-100 dark:bg-slate-700/80 text-slate-600 dark:text-slate-300 uppercase font-black sticky top-0">
                    <tr>
                      <th className="p-3">Fecha</th>
                      <th className="p-3">Invernadero</th>
                      <th className="p-3">Labor / Detalle</th>
                      <th className="p-3 text-right">Monto</th>
                      <th className="p-3 text-center">Acción</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y dark:divide-slate-700 font-bold text-slate-700 dark:text-slate-300">
                    {trabajadorDetalleEdicion.jornalesDetalle?.length === 0 ? (
                      <tr><td colSpan="5" className="p-6 text-center text-gray-400 italic">No hay registros individuales cargados</td></tr>
                    ) : (
                      trabajadorDetalleEdicion.jornalesDetalle?.map(j => (
                        <tr key={j.id} className="hover:bg-slate-50 dark:hover:bg-slate-700/50">
                          <td className="p-3">{j.fecha_labor}</td>
                          <td className="p-3 uppercase">{j.invernaderos?.nombre || 'General'}</td>
                          <td className="p-3 uppercase"><p>{j.tipo_labor}</p>{j.observaciones && <p className="text-[10px] text-gray-400 font-normal">{j.observaciones}</p>}</td>
                          <td className="p-3 text-right font-black text-emerald-700 dark:text-emerald-400">{formatoPesos(j.valor_pagar)}</td>
                          <td className="p-3 text-center">
                            <div className="flex gap-1 justify-center">
                              <button onClick={() => { prepararEdicionJornal(j); setTrabajadorDetalleEdicion(null); }} className="p-1.5 bg-amber-100 text-amber-700 dark:bg-amber-950/80 dark:text-amber-300 rounded-lg text-[10px] cursor-pointer" title="Editar">✏️</button>
                              <button onClick={async () => { await eliminarJornalPendiente(j.id); setTrabajadorDetalleEdicion(null); }} className="p-1.5 bg-red-100 text-red-700 dark:bg-red-950/80 dark:text-red-300 rounded-lg text-[10px] cursor-pointer" title="Eliminar">🗑️</button>
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
        </div>
      )}

    </div>
  );
}