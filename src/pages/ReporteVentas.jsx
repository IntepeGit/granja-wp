import React, { useState, useEffect } from 'react';
import { BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, Cell } from 'recharts';
import { supabase } from '../lib/supabase';
import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';

export default function ReporteVentas({ listaInvernaderos, datosDespachos, datosEgresos, datosPagos }) {
  const [invSeleccionado, setInvSeleccionado] = useState('');
  const [historicoNomina, setHistoricoNomina] = useState([]);
  const [historicoPagosNomina, setHistoricoPagosNomina] = useState([]);
  
  // ⚡ Estado para la vista SQL de ventas optimizadas desde Supabase
  const [ventasOptimizadasBD, setVentasOptimizadasBD] = useState([]);

  // Filtros de fecha
  const [fechaInicio, setFechaInicio] = useState('');
  const [fechaFin, setFechaFin] = useState('');
  const [filtroPeriodoRapido, setFiltroPeriodoRapido] = useState('historico_total');
  const [tabDetalleAccion, setTabDetalleAccion] = useState('ventas');

  useEffect(() => {
    cargarNominaCompleta();
    cargarVentasDesdeVistaSQL();
  }, [datosEgresos, datosDespachos]);

  const cargarVentasDesdeVistaSQL = async () => {
    try {
      const { data, error } = await supabase
        .from('reporte_ventas_totales')
        .select('*')
        .order('fecha_venta', { ascending: false });

      if (!error && data) {
        setVentasOptimizadasBD(data);
      }
    } catch (err) {
      console.error("Error consultando vista optimizada SQL, usando respaldo local:", err);
    }
  };

  const cargarNominaCompleta = async () => {
    try {
      const [resJornales, resPagosRealizados] = await Promise.all([
        supabase.from('nomina_jornales').select('*, nomina_trabajadores(nombre_completo), invernaderos(nombre)'),
        supabase.from('nomina_pagos_realizados').select('*, nomina_trabajadores(nombre_completo, tipo_pago)')
      ]);

      if (resJornales.data) setHistoricoNomina(resJornales.data);
      if (resPagosRealizados.data) setHistoricoPagosNomina(resPagosRealizados.data);
    } catch (err) {
      console.error("Error consultando nóminas desde Reporte:", err);
    }
  };

  const aplicarPeriodoRapido = (tipo) => {
    setFiltroPeriodoRapido(tipo);
    const ahora = new Date();
    const hoy = ahora.toISOString().split('T')[0];

    if (tipo === 'mes_actual') {
      setFechaInicio(new Date(ahora.getFullYear(), ahora.getMonth(), 1).toISOString().split('T')[0]);
      setFechaFin(hoy);
    } else if (tipo === 'ano_actual') {
      setFechaInicio(new Date(ahora.getFullYear(), 0, 1).toISOString().split('T')[0]);
      setFechaFin(hoy);
    } else if (tipo === 'historico_total') {
      setFechaInicio('');
      setFechaFin('');
    }
  };

  const formatoPesos = (valor) => new Intl.NumberFormat('es-CO', { 
    style: 'currency', currency: 'COP', minimumFractionDigits: 0 
  }).format(valor || 0);

  // --- 1. SEPARACIÓN DE INVERNADEROS Y NOMBRES ---
  const invernaderosActivos = (listaInvernaderos || []).filter(inv => inv.activo !== false);
  const idsInvernaderosActivos = invernaderosActivos.map(inv => inv.id?.toString());
  const numInvernaderosActivos = invernaderosActivos.length || 1;
  
  const objInvSeleccionado = (listaInvernaderos || []).find(i => i.id?.toString() === invSeleccionado);
  const nombreInvSeleccionado = objInvSeleccionado ? objInvSeleccionado.nombre?.toUpperCase() : 'TODOS LOS INVERNADEROS (EN PRODUCCIÓN)';

  const enRangoFecha = (fechaStr) => {
    if (!fechaStr) return true;
    if (!fechaInicio && !fechaFin) return true;
    if (fechaInicio && fechaStr < fechaInicio) return false;
    if (fechaFin && fechaStr > fechaFin) return false;
    return true;
  };

  // --- 2. FILTRADO COHERENTE CON PRORRATEO DE NÓMINA ---
  const fuenteVentas = ventasOptimizadasBD.length > 0 ? ventasOptimizadasBD : (datosDespachos || []);

  const despachosFiltrados = fuenteVentas.filter(d => {
    const coincideInv = !invSeleccionado 
      ? idsInvernaderosActivos.includes(d.invernadero_id?.toString()) 
      : d.invernadero_id?.toString() === invSeleccionado;
    return coincideInv && enRangoFecha(d.fecha_venta);
  });

  const egresosFiltrados = (datosEgresos || []).filter(g => {
    const coincideInv = !invSeleccionado 
      ? (!g.invernadero_id || idsInvernaderosActivos.includes(g.invernadero_id?.toString()))
      : g.invernadero_id?.toString() === invSeleccionado;
    return coincideInv && enRangoFecha(g.fecha_gasto);
  });

  const idsDespachosFiltrados = despachosFiltrados.map(d => d.id?.toString());
  const pagosFiltrados = (datosPagos || []).filter(p => {
    const perteneceADespacho = idsDespachosFiltrados.includes(p.despacho_id?.toString());
    return perteneceADespacho && enRangoFecha(p.fecha_pago);
  });

  // 👥 CÁLCULO Y FILTRADO AVANZADO DE MANO DE OBRA (CON QUINCENAS Y PRORRATEO)
  const nominaMapeada = (historicoPagosNomina || []).map(p => {
    const invNombrePago = (p.invernadero_nombre || '').toUpperCase();
    const esGeneral = !p.invernadero_nombre || invNombrePago.includes('GENERAL') || invNombrePago.includes('VARIOS');
    
    let coincide = false;
    let montoCalculado = parseFloat(p.monto_pagado || 0);
    let observacionProrrateo = '';

    if (!invSeleccionado) {
      coincide = true;
    } else {
      if (nombreInvSeleccionado && invNombrePago.includes(nombreInvSeleccionado)) {
        coincide = true;
      } else if (esGeneral) {
        coincide = true;
        montoCalculado = montoCalculado / numInvernaderosActivos;
        observacionProrrateo = ` (Prorrateado 1/${numInvernaderosActivos})`;
      }
    }

    return {
      ...p,
      coincide: coincide && enRangoFecha(p.fecha_pago),
      montoAplicado: montoCalculado,
      observacionProrrateo,
      tipoPagoTrabajador: p.nomina_trabajadores?.tipo_pago || 'Jornalero'
    };
  });

  const nominaFiltrada = nominaMapeada.filter(n => n.coincide);

  // --- 3. CÁLCULOS CONSOLIDADOS Y SEPARACIÓN DE CONCEPTOS ---
  const totalVentas = despachosFiltrados.reduce((acc, d) => acc + parseFloat(d.total_venta || 0), 0);
  
  const totalInsumosGastos = egresosFiltrados
    .filter(g => g.categoria !== 'Mano de obra' && g.categoria !== 'Quincena')
    .reduce((acc, g) => acc + parseFloat(g.monto || 0), 0);

  // Discriminación de Mano de Obra: Jornales vs Quincenas
  const totalJornalesSabado = nominaFiltrada
    .filter(n => (n.tipoPagoTrabajador || '').includes('Sábado') || (n.tipoPagoTrabajador || '').includes('Jornalero'))
    .reduce((acc, n) => acc + n.montoAplicado, 0);

  const totalQuincenasFijas = nominaFiltrada
    .filter(n => (n.tipoPagoTrabajador || '').includes('Quincenal') || (n.tipoPagoTrabajador || '').includes('Fijo'))
    .reduce((acc, n) => acc + n.montoAplicado, 0);

  const totalManoObra = totalJornalesSabado + totalQuincenasFijas;
  const totalGastos = totalInsumosGastos + totalManoObra;

  const pagosRecibidos = pagosFiltrados.reduce((acc, p) => acc + parseFloat(p.monto || 0), 0);
  const utilidadNeta = totalVentas - totalGastos;
  const margen = totalVentas > 0 ? ((utilidadNeta / totalVentas) * 100).toFixed(1) : 0;
  const cuentasPorCobrar = totalVentas - pagosRecibidos;

  const dataGrafica = [
    { name: 'Gastos', valor: totalGastos, color: '#ef4444' }, // red-500
    { name: 'Ventas', valor: totalVentas, color: '#10b981' }, // emerald-500
    { name: 'Utilidad', valor: utilidadNeta, color: '#3b82f6' } // blue-500
  ];

  // --- 📊 EXPORTAR A EXCEL ---
  const exportarReporteAExcel = async () => {
    try {
      const workbook = new ExcelJS.Workbook();
      const hoyStr = new Date().toISOString().split('T')[0];

      // HOJA 1: RESUMEN FINANCIERO DESGLOSADO
      const wsResumen = workbook.addWorksheet('Resumen Financiero');
      wsResumen.columns = [
        { header: 'CONCEPTO FINANCIERO', key: 'concepto', width: 45 },
        { header: 'VALOR TOTAL (COP)', key: 'valor', width: 25 }
      ];

      const filaVentas = wsResumen.addRow({ concepto: 'VENTAS TOTALES (INGRESOS)', valor: totalVentas });
      const filaInsumos = wsResumen.addRow({ concepto: 'GASTOS DE INSUMOS Y MATERIALES', valor: totalInsumosGastos });
      const filaJornales = wsResumen.addRow({ concepto: 'MANO DE OBRA - JORNALES (SÁBADO)', valor: totalJornalesSabado });
      const filaQuincenas = wsResumen.addRow({ concepto: 'MANO DE OBRA - SUELDOS FIJOS (QUINCENAL)', valor: totalQuincenasFijas });
      const filaTotalManoObra = wsResumen.addRow({ concepto: 'TOTAL MANO DE OBRA ACUMULADA', valor: totalManoObra });
      const filaGastosTotales = wsResumen.addRow({ concepto: 'GASTOS TOTALES (INSUMOS + MANO DE OBRA)', valor: totalGastos });
      const filaUtilidad = wsResumen.addRow({ concepto: 'UTILIDAD NETA OPERACIONAL', valor: utilidadNeta });
      const filaMargen = wsResumen.addRow({ concepto: 'MARGEN DE RENDIMIENTO', valor: totalVentas > 0 ? (utilidadNeta / totalVentas) : 0 });
      const filaRecaudado = wsResumen.addRow({ concepto: 'TOTAL RECAUDADO (CAJA DE COBRO)', valor: pagosRecibidos });
      const filaCartera = wsResumen.addRow({ concepto: 'CUENTAS POR COBRAR (CARTERA PENDIENTE)', valor: cuentasPorCobrar });

      // Formato Encabezado Hoja 1
      const headerRow1 = wsResumen.getRow(1);
      headerRow1.height = 24;
      headerRow1.eachCell(c => {
        c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E293B' } };
        c.font = { name: 'Arial', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
        c.alignment = { vertical: 'middle', horizontal: 'center' };
      });

      // Estilos para Filas
      [
        filaVentas, filaInsumos, filaJornales, filaQuincenas, 
        filaTotalManoObra, filaGastosTotales, filaUtilidad, 
        filaMargen, filaRecaudado, filaCartera
      ].forEach(row => {
        row.height = 22;
        row.getCell(1).font = { name: 'Arial', size: 10, bold: true };
        row.getCell(2).font = { name: 'Arial', size: 11, bold: true };
      });

      // Colores de Enfasis
      filaVentas.getCell(2).font.color = { argb: 'FF15803D' };
      filaInsumos.getCell(2).font.color = { argb: 'FFB91C1C' };
      filaJornales.getCell(2).font.color = { argb: 'FFC2410C' };
      filaQuincenas.getCell(2).font.color = { argb: 'FF6B21A8' };
      filaTotalManoObra.getCell(2).font.color = { argb: 'FF0284C7' };
      filaGastosTotales.getCell(2).font.color = { argb: 'FFB91C1C' };
      filaUtilidad.getCell(2).font.color = { argb: 'FF1D4ED8' };

      [
        filaVentas, filaInsumos, filaJornales, filaQuincenas, 
        filaTotalManoObra, filaGastosTotales, filaUtilidad, 
        filaRecaudado, filaCartera
      ].forEach(row => {
        row.getCell(2).numFmt = '"$"#,##0';
        row.getCell(2).alignment = { horizontal: 'right', vertical: 'middle' };
      });

      filaMargen.getCell(2).numFmt = '0.0%';
      filaMargen.getCell(2).alignment = { horizontal: 'right', vertical: 'middle' };

      // HOJA 2: DETALLE DE VENTAS
      const wsVentas = workbook.addWorksheet('Detalle Ventas');
      wsVentas.columns = [
        { header: 'FECHA VENTA', key: 'fecha', width: 15 },
        { header: 'REMISIÓN N°', key: 'remision', width: 15 },
        { header: 'INVERNADERO', key: 'inv', width: 22 },
        { header: 'CLIENTE', key: 'cliente', width: 25 },
        { header: 'VALOR DESPACHO', key: 'total', width: 22 }
      ];

      despachosFiltrados.forEach(d => {
        wsVentas.addRow({
          fecha: d.fecha_venta || '',
          remision: d.numero_remision || 'S/N',
          inv: (d.invernadero_nombre || d.invernaderos?.nombre || 'General').toUpperCase(),
          cliente: (d.cliente_nombre || d.clientes?.nombre_completo || 'Particular').toUpperCase(),
          total: parseFloat(d.total_venta || 0)
        });
      });

      const ultFilaV = wsVentas.rowCount;
      const filaTotV = wsVentas.addRow({
        cliente: 'TOTAL GENERAL VENTAS:',
        total: { formula: `=SUM(E2:E${ultFilaV})` }
      });
      filaTotV.getCell('cliente').font = { name: 'Arial', size: 10, bold: true };
      filaTotV.getCell('total').font = { name: 'Arial', size: 11, bold: true, color: { argb: 'FF15803D' } };

      // HOJA 3: DETALLE GASTOS INSUMOS
      const wsGastos = workbook.addWorksheet('Detalle Gastos Insumos');
      wsGastos.columns = [
        { header: 'FECHA GASTO', key: 'fecha', width: 15 },
        { header: 'COMPROBANTE', key: 'doc', width: 18 },
        { header: 'INVERNADERO', key: 'inv', width: 18 },
        { header: 'CATEGORÍA', key: 'cat', width: 20 },
        { header: 'CONCEPTO / DETALLE', key: 'desc', width: 35 },
        { header: 'MONTO GASTO', key: 'monto', width: 20 }
      ];

      egresosFiltrados.filter(g => g.categoria !== 'Mano de obra' && g.categoria !== 'Quincena').forEach(g => {
        wsGastos.addRow({
          fecha: g.fecha_gasto || '',
          doc: g.numero_comprobante || 'S/N',
          inv: (g.invernaderos?.nombre || 'General').toUpperCase(),
          cat: (g.categoria || 'Varios').toUpperCase(),
          desc: (g.descripcion || '').toUpperCase(),
          monto: parseFloat(g.monto || 0)
        });
      });

      const ultFilaG = wsGastos.rowCount;
      const filaTotG = wsGastos.addRow({
        desc: 'TOTAL GENERAL INSUMOS:',
        monto: { formula: `=SUM(F2:F${ultFilaG})` }
      });
      filaTotG.getCell('desc').font = { name: 'Arial', size: 10, bold: true };
      filaTotG.getCell('monto').font = { name: 'Arial', size: 11, bold: true, color: { argb: 'FFB91C1C' } };

      // HOJA 4: DETALLE MANO DE OBRA (COMPLETA CON QUINCENAS Y JORNALES)
      const wsManoObra = workbook.addWorksheet('Detalle Mano de Obra');
      wsManoObra.columns = [
        { header: 'COMP. N°', key: 'comp', width: 14 },
        { header: 'FECHA PAGO', key: 'fecha', width: 15 },
        { header: 'INVERNADERO', key: 'inv', width: 25 },
        { header: 'COLABORADOR', key: 'nombre', width: 30 },
        { header: 'MODALIDAD PAGO', key: 'tipo', width: 22 },
        { header: 'NETO ASIGNADO / PAGADO', key: 'monto', width: 24 }
      ];

      nominaFiltrada.forEach(n => {
        wsManoObra.addRow({
          comp: `NOM-${String(n.id).padStart(4, '0')}`,
          fecha: n.fecha_pago || '',
          inv: `${(n.invernadero_nombre || 'GENERAL / VARIOS').toUpperCase()}${n.observacionProrrateo}`,
          nombre: (n.nomina_trabajadores?.nombre_completo || 'OPERARIO').toUpperCase(),
          tipo: (n.tipoPagoTrabajador || 'Jornalero').toUpperCase(),
          monto: parseFloat(n.montoAplicado || 0)
        });
      });

      const ultFilaM = wsManoObra.rowCount;
      const filaTotM = wsManoObra.addRow({
        tipo: 'TOTAL MANO DE OBRA:',
        monto: { formula: `=SUM(F2:F${ultFilaM})` }
      });
      filaTotM.getCell('tipo').font = { name: 'Arial', size: 10, bold: true };
      filaTotM.getCell('monto').font = { name: 'Arial', size: 11, bold: true, color: { argb: 'FF6B21A8' } };

      // Formateo visual uniforme
      [wsVentas, wsGastos, wsManoObra].forEach(ws => {
        const head = ws.getRow(1);
        head.height = 24;
        head.eachCell(c => {
          c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF117097' } };
          c.font = { name: 'Arial', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
          c.alignment = { vertical: 'middle', horizontal: 'center' };
        });

        const totalRowsCount = ws.rowCount;

        ws.eachRow((row, idx) => {
          if (idx === 1) return;
          row.height = 20;

          if (idx === totalRowsCount) {
            row.eachCell((cell) => {
              cell.border = {
                top: { style: 'thin', color: { argb: 'FF000000' } },
                bottom: { style: 'double', color: { argb: 'FF000000' } }
              };
            });
            row.getCell(ws.columnCount).numFmt = '"$"#,##0';
            row.getCell(ws.columnCount).alignment = { horizontal: 'right', vertical: 'middle' };
            return;
          }

          const esCebra = idx % 2 === 0;
          row.eachCell((cell, colNum) => {
            cell.font = { name: 'Arial', size: 9 };
            if (esCebra) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEBF5FB' } };

            if (colNum === ws.columnCount) {
              cell.numFmt = '"$"#,##0';
              cell.alignment = { horizontal: 'right', vertical: 'middle' };
            } else {
              cell.alignment = { horizontal: 'left', vertical: 'middle' };
            }
          });
        });

        ws.autoFilter = { from: { row: 1, column: 1 }, to: { row: totalRowsCount - 1, column: ws.columnCount } };
      });

      const buffer = await workbook.xlsx.writeBuffer();
      const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      saveAs(blob, `REPORTE_VENTAS_${nombreInvSeleccionado.replace(/ /g, '_')}_${hoyStr}.xlsx`);

    } catch (err) {
      console.error("Error al exportar reporte de ventas:", err);
    }
  };

  return (
    <div className="space-y-6 pb-20 text-slate-800 dark:text-slate-200 font-sans transition-colors duration-300">
      
      {/* 🚀 CABECERA PRINCIPAL MODERNA */}
      <div className="bg-white dark:bg-slate-800 p-6 rounded-3xl shadow-xl border border-slate-200 dark:border-slate-700 flex flex-col md:flex-row justify-between items-center gap-4 transition-colors duration-300">
        <div className="space-y-1">
          <div className="flex items-center gap-3">
            <span className="text-2xl p-2 bg-emerald-700/10 dark:bg-emerald-500/20 rounded-xl text-emerald-700 dark:text-emerald-400">📈</span>
            <h2 className="text-xl font-black tracking-tight text-slate-900 dark:text-white uppercase">Reporte de Ventas y Finanzas</h2>
          </div>
          <p className="text-xs text-slate-500 dark:text-slate-400 font-medium">Análisis de rentabilidad, ingresos, gastos y estado de cartera.</p>
        </div>
        
        <div className="flex items-center gap-3 w-full md:w-auto">
          <button onClick={exportarReporteAExcel} className="flex-1 md:flex-none px-5 py-2.5 bg-emerald-600 hover:bg-emerald-700 text-white font-black text-xs uppercase tracking-wider rounded-xl shadow-md transition-all flex items-center justify-center gap-2 cursor-pointer">
            <span>📊</span> Exportar a Excel (Reporte Total)
          </button>
        </div>
      </div>

      {/* 🔍 SECCIÓN DE FILTROS SUPERIORES */}
      <div className="bg-white dark:bg-slate-800 p-5 rounded-2xl shadow-md border border-slate-200 dark:border-slate-700 grid grid-cols-1 md:grid-cols-3 gap-4 items-end transition-colors duration-300">
        
        <div>
          <label className="text-[10px] font-black text-gray-400 dark:text-slate-400 uppercase tracking-widest mb-1 block">Análisis de Invernadero / Bloque</label>
          <select 
            className="w-full border border-slate-200 dark:border-slate-700 p-2.5 rounded-xl font-black text-slate-800 dark:text-white bg-slate-50 dark:bg-slate-900 outline-none focus:border-[#117097] text-xs transition-colors"
            value={invSeleccionado}
            onChange={(e) => setInvSeleccionado(e.target.value)}
          >
            <option value="">-- TODOS LOS INVERNADEROS (EN PRODUCCIÓN) --</option>
            <optgroup label="🌱 EN PRODUCCIÓN (OPERATIVOS)">
              {(listaInvernaderos || []).filter(i => i.activo !== false).map(inv => (
                <option key={inv.id} value={inv.id}>{inv.nombre?.toUpperCase()}</option>
              ))}
            </optgroup>
            <optgroup label="📁 HISTÓRICO / ARCHIVADOS">
              {(listaInvernaderos || []).filter(i => i.activo === false).map(inv => (
                <option key={inv.id} value={inv.id}>{inv.nombre?.toUpperCase()} (ARCHIVADO)</option>
              ))}
            </optgroup>
          </select>
        </div>

        <div className="md:col-span-2 flex flex-col md:flex-row gap-4 justify-between items-end">
          <div className="flex-1 space-y-1 w-full">
            <label className="text-[10px] font-black text-gray-400 dark:text-slate-400 uppercase tracking-widest mb-1 block">Rango de Fechas Personalizado</label>
            <div className="flex items-center gap-2">
              <input type="date" className="w-full bg-slate-50 dark:bg-slate-900 border border-slate-200 dark:border-slate-700 text-slate-800 dark:text-white p-2.5 rounded-xl font-bold text-xs outline-none focus:border-[#117097]" value={fechaInicio} onChange={e => { setFechaInicio(e.target.value); setFiltroPeriodoRapido('personalizado'); }} />
              <span className="text-slate-300 dark:text-slate-600 font-black">-</span>
              <input type="date" className="w-full bg-slate-50 dark:bg-slate-900 border border-slate-200 dark:border-slate-700 text-slate-800 dark:text-white p-2.5 rounded-xl font-bold text-xs outline-none focus:border-[#117097]" value={fechaFin} onChange={e => { setFechaFin(e.target.value); setFiltroPeriodoRapido('personalizado'); }} />
            </div>
          </div>
          
          <div className="flex gap-2 w-full md:w-auto overflow-x-auto pb-1 md:pb-0">
            <button onClick={() => aplicarPeriodoRapido('mes_actual')} className={`px-4 py-2.5 rounded-xl text-[10px] font-black uppercase transition-all whitespace-nowrap cursor-pointer ${filtroPeriodoRapido === 'mes_actual' ? 'bg-[#117097] text-white shadow' : 'bg-slate-100 dark:bg-slate-700 text-slate-500 dark:text-slate-300 hover:bg-slate-200 dark:hover:bg-slate-600'}`}>📅 Mes Actual</button>
            <button onClick={() => aplicarPeriodoRapido('ano_actual')} className={`px-4 py-2.5 rounded-xl text-[10px] font-black uppercase transition-all whitespace-nowrap cursor-pointer ${filtroPeriodoRapido === 'ano_actual' ? 'bg-[#117097] text-white shadow' : 'bg-slate-100 dark:bg-slate-700 text-slate-500 dark:text-slate-300 hover:bg-slate-200 dark:hover:bg-slate-600'}`}>🗓️ Año 2026</button>
            <button onClick={() => aplicarPeriodoRapido('historico_total')} className={`px-4 py-2.5 rounded-xl text-[10px] font-black uppercase transition-all whitespace-nowrap cursor-pointer ${filtroPeriodoRapido === 'historico_total' ? 'bg-[#117097] text-white shadow' : 'bg-slate-100 dark:bg-slate-700 text-slate-500 dark:text-slate-300 hover:bg-slate-200 dark:hover:bg-slate-600'}`}>🌐 Histórico Todo</button>
          </div>
        </div>

      </div>

      {/* 📈 CARDS DINÁMICAS (KPIs) */}
      <div className="grid grid-cols-1 sm:grid-cols-2 xl:grid-cols-4 gap-4">
        <div className="bg-white dark:bg-slate-800 p-6 rounded-3xl shadow-xl border-b-4 border-red-500 text-center flex flex-col justify-center transition-transform hover:-translate-y-1">
          <p className="text-[10px] font-black text-slate-400 dark:text-slate-400 uppercase tracking-widest">Gastos Totales (Insumos + Obra)</p>
          <p className="text-2xl font-black text-red-600 dark:text-red-400 mt-2">{formatoPesos(totalGastos)}</p>
        </div>
        <div className="bg-white dark:bg-slate-800 p-6 rounded-3xl shadow-xl border-b-4 border-emerald-500 text-center flex flex-col justify-center transition-transform hover:-translate-y-1">
          <p className="text-[10px] font-black text-slate-400 dark:text-slate-400 uppercase tracking-widest">Ventas Totales (Ingresos)</p>
          <p className="text-2xl font-black text-emerald-600 dark:text-emerald-400 mt-2">{formatoPesos(totalVentas)}</p>
        </div>
        <div className="bg-white dark:bg-slate-800 p-6 rounded-3xl shadow-xl border-b-4 border-blue-500 text-center flex flex-col justify-center transition-transform hover:-translate-y-1">
          <p className="text-[10px] font-black text-slate-400 dark:text-slate-400 uppercase tracking-widest">Utilidad Neta</p>
          <p className="text-2xl font-black text-blue-600 dark:text-blue-400 mt-2">{formatoPesos(utilidadNeta)}</p>
        </div>
        <div className="bg-slate-800 dark:bg-slate-900 p-6 rounded-3xl shadow-xl text-center flex flex-col justify-center border-b-4 border-slate-900 transition-transform hover:-translate-y-1">
          <p className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Margen de Rendimiento</p>
          <p className="text-3xl font-black text-white mt-1">{margen}%</p>
        </div>
      </div>

      {/* 📊 GRÁFICA Y CARTERA */}
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
        <div className="bg-white dark:bg-slate-800 p-6 rounded-3xl shadow-xl h-72 border border-slate-200 dark:border-slate-700 flex flex-col transition-colors duration-300">
          <p className="text-[10px] font-black text-slate-400 uppercase mb-4 tracking-widest">Comparativa Rendimiento Financiero</p>
          <div className="flex-1 min-h-0">
            <ResponsiveContainer width="100%" height="100%">
              <BarChart data={dataGrafica} margin={{ top: 10, right: 10, left: -20, bottom: 0 }}>
                <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#334155" opacity={0.2} />
                <XAxis dataKey="name" axisLine={false} tickLine={false} tick={{fontSize: 10, fontStyle: 'bold', fill: '#94a3b8'}} />
                <YAxis hide />
                <Tooltip cursor={{fill: 'transparent'}} formatter={(value) => formatoPesos(value)} contentStyle={{borderRadius: '16px', border: '1px solid #334155', backgroundColor: '#1e293b', color: '#fff', boxShadow: '0 10px 15px -3px rgb(0 0 0 / 0.1)', fontWeight: 'bold'}} />
                <Bar dataKey="valor" radius={[6, 6, 0, 0]} maxBarSize={60}>
                  {dataGrafica.map((entry, index) => (
                    <Cell key={`cell-${index}`} fill={entry.color} />
                  ))}
                </Bar>
              </BarChart>
            </ResponsiveContainer>
          </div>
        </div>

        <div className="bg-white dark:bg-slate-800 p-6 rounded-3xl shadow-xl flex flex-col justify-center h-72 border border-slate-200 dark:border-slate-700 transition-colors duration-300">
          <p className="text-[10px] font-black text-slate-400 uppercase mb-6 tracking-widest">Estado de Recaudos y Cartera</p>
          <div className="space-y-4">
            <div className="flex justify-between items-center bg-emerald-50/50 dark:bg-emerald-950/40 p-5 rounded-2xl border border-emerald-100 dark:border-emerald-900 shadow-sm">
              <span className="text-[10px] font-black text-emerald-800 dark:text-emerald-400 uppercase tracking-widest flex items-center gap-2"><span className="text-lg">💰</span> Total Cobrado (Caja Recaudada)</span>
              <span className="font-black text-emerald-700 dark:text-emerald-300 text-lg">{formatoPesos(pagosRecibidos)}</span>
            </div>
            <div className="flex justify-between items-center bg-amber-50/50 dark:bg-amber-950/40 p-5 rounded-2xl border border-amber-100 dark:border-amber-900 shadow-sm">
              <span className="text-[10px] font-black text-amber-800 dark:text-amber-400 uppercase tracking-widest flex items-center gap-2"><span className="text-lg">⏳</span> Por Cobrar (Cartera Pendiente)</span>
              <span className="font-black text-amber-700 dark:text-amber-300 text-lg">{formatoPesos(cuentasPorCobrar)}</span>
            </div>
          </div>
        </div>
      </div>

      {/* 📁 VISTA EXPLORATORIA EN PANTALLA (ANCHO COMPLETO) */}
      <div className="bg-white dark:bg-slate-800 rounded-3xl shadow-xl overflow-hidden border border-slate-200 dark:border-slate-700 transition-colors duration-300">
        
        <div className="p-4 bg-slate-800 dark:bg-slate-900 text-white font-black text-xs uppercase tracking-wider flex flex-col md:flex-row justify-between items-center gap-4">
          <span className="truncate w-full md:w-auto">🔍 Explorador: {nombreInvSeleccionado}</span>
          <div className="flex gap-2 overflow-x-auto w-full md:w-auto pb-1 md:pb-0">
            <button onClick={() => setTabDetalleAccion('ventas')} className={`px-4 py-2 rounded-xl text-[10px] uppercase font-black transition-all whitespace-nowrap cursor-pointer ${tabDetalleAccion === 'ventas' ? 'bg-emerald-600 text-white shadow' : 'bg-slate-700 text-slate-300 hover:bg-slate-600'}`}>🛒 Ventas ({despachosFiltrados.length})</button>
            <button onClick={() => setTabDetalleAccion('gastos')} className={`px-4 py-2 rounded-xl text-[10px] uppercase font-black transition-all whitespace-nowrap cursor-pointer ${tabDetalleAccion === 'gastos' ? 'bg-red-600 text-white shadow' : 'bg-slate-700 text-slate-300 hover:bg-slate-600'}`}>💸 Insumos ({egresosFiltrados.length})</button>
            <button onClick={() => setTabDetalleAccion('nomina')} className={`px-4 py-2 rounded-xl text-[10px] uppercase font-black transition-all whitespace-nowrap cursor-pointer ${tabDetalleAccion === 'nomina' ? 'bg-purple-600 text-white shadow' : 'bg-slate-700 text-slate-300 hover:bg-slate-600'}`}>👥 Mano Obra ({nominaFiltrada.length})</button>
          </div>
        </div>

        <div className="overflow-x-auto max-h-[400px] overflow-y-auto">
          {tabDetalleAccion === 'ventas' && (
            <table className="w-full text-left border-collapse text-xs">
              <thead>
                <tr className="bg-slate-100 dark:bg-slate-700/80 text-slate-600 dark:text-slate-300 uppercase font-black text-[10px] tracking-wider border-b border-slate-200 dark:border-slate-700 sticky top-0 z-10">
                  <th className="p-4">Fecha</th>
                  <th className="p-4">Remisión N°</th>
                  <th className="p-4">Invernadero</th>
                  <th className="p-4">Cliente</th>
                  <th className="p-4 text-right">Monto Venta</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100 dark:divide-slate-700/60 font-bold text-slate-700 dark:text-slate-300">
                {despachosFiltrados.length === 0 ? (
                  <tr><td colSpan="5" className="p-8 text-center text-slate-400 italic">No hay registros de ventas para este filtro.</td></tr>
                ) : (
                  despachosFiltrados.map((d, index) => (
                    <tr key={d.id} className={`${index % 2 === 0 ? 'bg-white dark:bg-slate-800' : 'bg-slate-50/50 dark:bg-slate-800/60'} hover:bg-emerald-50/50 dark:hover:bg-slate-700/50 transition-colors border-l-4 border-emerald-500`}>
                      <td className="p-4 whitespace-nowrap">{d.fecha_venta}</td>
                      <td className="p-4 font-black text-slate-900 dark:text-white">{d.numero_remision || 'S/N'}</td>
                      <td className="p-4 uppercase text-[10px]">{d.invernadero_nombre || d.invernaderos?.nombre || 'General'}</td>
                      <td className="p-4 uppercase">{d.cliente_nombre || d.clientes?.nombre_completo || 'Particular'}</td>
                      <td className="p-4 text-right text-emerald-700 dark:text-emerald-400 font-black text-sm">{formatoPesos(d.total_venta)}</td>
                    </tr>
                  ))
                )}
              </tbody>
            </table>
          )}

          {tabDetalleAccion === 'gastos' && (
            <table className="w-full text-left border-collapse text-xs">
              <thead>
                <tr className="bg-slate-100 dark:bg-slate-700/80 text-slate-600 dark:text-slate-300 uppercase font-black text-[10px] tracking-wider border-b border-slate-200 dark:border-slate-700 sticky top-0 z-10">
                  <th className="p-4">Fecha</th>
                  <th className="p-4">Comprobante</th>
                  <th className="p-4">Categoría</th>
                  <th className="p-4">Descripción</th>
                  <th className="p-4 text-right">Monto Gasto</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100 dark:divide-slate-700/60 font-bold text-slate-700 dark:text-slate-300">
                {egresosFiltrados.length === 0 ? (
                  <tr><td colSpan="5" className="p-8 text-center text-slate-400 italic">No hay gastos de insumos para este filtro.</td></tr>
                ) : (
                  egresosFiltrados.map((g, index) => (
                    <tr key={g.id} className={`${index % 2 === 0 ? 'bg-white dark:bg-slate-800' : 'bg-slate-50/50 dark:bg-slate-800/60'} hover:bg-red-50/50 dark:hover:bg-slate-700/50 transition-colors border-l-4 border-red-500`}>
                      <td className="p-4 whitespace-nowrap">{g.fecha_gasto}</td>
                      <td className="p-4 font-black text-slate-900 dark:text-white">{g.numero_comprobante || 'S/N'}</td>
                      <td className="p-4 uppercase text-[10px] text-[#117097] dark:text-sky-400">{g.categoria}</td>
                      <td className="p-4 uppercase text-slate-500 dark:text-slate-400 text-[10px]">{g.descripcion}</td>
                      <td className="p-4 text-right text-red-600 dark:text-red-400 font-black text-sm">{formatoPesos(g.monto)}</td>
                    </tr>
                  ))
                )}
              </tbody>
            </table>
          )}

          {tabDetalleAccion === 'nomina' && (
            <table className="w-full text-left border-collapse text-xs">
              <thead>
                <tr className="bg-slate-100 dark:bg-slate-700/80 text-slate-600 dark:text-slate-300 uppercase font-black text-[10px] tracking-wider border-b border-slate-200 dark:border-slate-700 sticky top-0 z-10">
                  <th className="p-4 text-center">Comp. N°</th>
                  <th className="p-4">Fecha Pago</th>
                  <th className="p-4">Invernadero / Asignación</th>
                  <th className="p-4">Trabajador</th>
                  <th className="p-4">Modalidad</th>
                  <th className="p-4 text-right">Neto Aplicado</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100 dark:divide-slate-700/60 font-bold text-slate-700 dark:text-slate-300">
                {nominaFiltrada.length === 0 ? (
                  <tr><td colSpan="6" className="p-8 text-center text-slate-400 italic">No hay pagos de nómina para este filtro.</td></tr>
                ) : (
                  nominaFiltrada.map((n, index) => (
                    <tr key={n.id} className={`${index % 2 === 0 ? 'bg-white dark:bg-slate-800' : 'bg-slate-50/50 dark:bg-slate-800/60'} hover:bg-purple-50/50 dark:hover:bg-slate-700/50 transition-colors border-l-4 border-purple-600`}>
                      <td className="p-4 text-center font-black text-slate-900 dark:text-white">NOM-{String(n.id).padStart(4, '0')}</td>
                      <td className="p-4 whitespace-nowrap">{n.fecha_pago}</td>
                      <td className="p-4 uppercase font-black text-purple-700 dark:text-purple-400 text-[10px]">
                        {n.invernadero_nombre || 'GENERAL / VARIOS'}{n.observacionProrrateo}
                      </td>
                      <td className="p-4 uppercase">{n.nomina_trabajadores?.nombre_completo || 'OPERARIO'}</td>
                      <td className="p-4 uppercase text-slate-500 dark:text-slate-400 text-[10px]">{n.tipoPagoTrabajador}</td>
                      <td className="p-4 text-right text-purple-700 dark:text-purple-400 font-black text-sm">{formatoPesos(n.montoAplicado)}</td>
                    </tr>
                  ))
                )}
              </tbody>
            </table>
          )}
        </div>
      </div>

    </div>
  );
}