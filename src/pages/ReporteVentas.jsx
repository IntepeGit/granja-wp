import React, { useState, useEffect } from 'react';
import { BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, Cell } from 'recharts';
import { supabase } from '../lib/supabase';
import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';

export default function ReporteVentas({ listaInvernaderos, datosDespachos, datosEgresos, datosPagos }) {
  const [invSeleccionado, setInvSeleccionado] = useState('');
  const [historicoNomina, setHistoricoNomina] = useState([]);
  const [historicoPagosNomina, setHistoricoPagosNomina] = useState([]);

  // Por defecto inicializa en 'historico_total' para coincidir con la vista completa
  const [fechaInicio, setFechaInicio] = useState('');
  const [fechaFin, setFechaFin] = useState('');
  const [filtroPeriodoRapido, setFiltroPeriodoRapido] = useState('historico_total');
  const [tabDetalleAccion, setTabDetalleAccion] = useState('ventas');

  useEffect(() => {
    cargarNominaCompleta();
  }, [datosEgresos]);

  const cargarNominaCompleta = async () => {
    try {
      const [resJornales, resPagosRealizados] = await Promise.all([
        supabase.from('nomina_jornales').select('*, nomina_trabajadores(nombre_completo), invernaderos(nombre)'),
        supabase.from('nomina_pagos_realizados').select('*, nomina_trabajadores(nombre_completo)')
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
  
  const objInvSeleccionado = (listaInvernaderos || []).find(i => i.id?.toString() === invSeleccionado);
  const nombreInvSeleccionado = objInvSeleccionado ? objInvSeleccionado.nombre?.toUpperCase() : 'TODOS LOS INVERNADEROS (EN PRODUCCIÓN)';

  const enRangoFecha = (fechaStr) => {
    if (!fechaStr) return true;
    if (!fechaInicio && !fechaFin) return true;
    if (fechaInicio && fechaStr < fechaInicio) return false;
    if (fechaFin && fechaStr > fechaFin) return false;
    return true;
  };

  // --- 2. FILTRADO CONHERENTE SEGÚN LA SELECCIÓN ---
  // A. Despachos / Ventas
  const despachosFiltrados = (datosDespachos || []).filter(d => {
    const coincideInv = !invSeleccionado 
      ? idsInvernaderosActivos.includes(d.invernadero_id?.toString()) 
      : d.invernadero_id?.toString() === invSeleccionado;
    return coincideInv && enRangoFecha(d.fecha_venta);
  });

  // B. Gastos e Insumos (Filtrado estricto por lotes activos si no hay uno seleccionado)
  const egresosFiltrados = (datosEgresos || []).filter(g => {
    const coincideInv = !invSeleccionado 
      ? (!g.invernadero_id || idsInvernaderosActivos.includes(g.invernadero_id?.toString()))
      : g.invernadero_id?.toString() === invSeleccionado;
    return coincideInv && enRangoFecha(g.fecha_gasto);
  });

  // C. Recaudos
  const idsDespachosFiltrados = despachosFiltrados.map(d => d.id?.toString());
  const pagosFiltrados = (datosPagos || []).filter(p => {
    const perteneceADespacho = idsDespachosFiltrados.includes(p.despacho_id?.toString());
    return perteneceADespacho && enRangoFecha(p.fecha_pago);
  });

  // D. Nómina (Mano de Obra)
  const nominaFiltrada = (historicoPagosNomina || []).filter(p => {
    const coincideInv = !invSeleccionado
      ? (!p.invernadero_nombre || invernaderosActivos.some(i => i.nombre?.toUpperCase() === p.invernadero_nombre?.toUpperCase()))
      : (p.invernadero_nombre || '').toUpperCase().includes(nombreInvSeleccionado.toUpperCase());
    return coincideInv && enRangoFecha(p.fecha_pago);
  });

  // --- 3. CÁLCULOS CONSOLIDADOS ---
  const totalVentas = despachosFiltrados.reduce((acc, d) => acc + parseFloat(d.total_venta || 0), 0);
  
  const totalInsumosGastos = egresosFiltrados
    .filter(g => g.categoria !== 'Mano de obra' && g.categoria !== 'Quincena')
    .reduce((acc, g) => acc + parseFloat(g.monto || 0), 0);

  const totalManoObra = nominaFiltrada.reduce((acc, n) => acc + parseFloat(n.monto_pagado || 0), 0);
  const totalGastos = totalInsumosGastos + totalManoObra;

  const pagosRecibidos = pagosFiltrados.reduce((acc, p) => acc + parseFloat(p.monto || 0), 0);
  const utilidadNeta = totalVentas - totalGastos;
  const margen = totalVentas > 0 ? ((utilidadNeta / totalVentas) * 100).toFixed(1) : 0;
  const cuentasPorCobrar = totalVentas - pagosRecibidos;

  const dataGrafica = [
    { name: 'Gastos', valor: totalGastos, color: '#ef4444' },
    { name: 'Ventas', valor: totalVentas, color: '#22c55e' },
    { name: 'Utilidad', valor: utilidadNeta, color: '#3b82f6' }
  ];

  // --- 📊 EXPORTAR A EXCEL ---
  const exportarReporteAExcel = async () => {
    try {
      const workbook = new ExcelJS.Workbook();
      const hoyStr = new Date().toISOString().split('T')[0];

      // HOJA 1: RESUMEN FINANCIERO
      const wsResumen = workbook.addWorksheet('Resumen Financiero');
      wsResumen.columns = [
        { header: 'CONCEPTO FINANCIERO', key: 'concepto', width: 42 },
        { header: 'VALOR TOTAL (COP)', key: 'valor', width: 25 }
      ];

      const filaVentas = wsResumen.addRow({ concepto: 'VENTAS TOTALES (INGRESOS)', valor: totalVentas });
      const filaGastos = wsResumen.addRow({ concepto: 'GASTOS TOTALES (CON MANO DE OBRA)', valor: totalGastos });
      const filaUtilidad = wsResumen.addRow({ concepto: 'UTILIDAD NETA OPERACIONAL', valor: utilidadNeta });
      const filaMargen = wsResumen.addRow({ concepto: 'MARGEN DE RENDIMIENTO', valor: totalVentas > 0 ? (utilidadNeta / totalVentas) : 0 });
      const filaRecaudado = wsResumen.addRow({ concepto: 'TOTAL RECAUDADO (CAJA DE COBRO)', valor: pagosRecibidos });
      const filaCartera = wsResumen.addRow({ concepto: 'CUENTAS POR COBRAR (CARTERA PENDIENTE)', valor: cuentasPorCobrar });

      const headerRow1 = wsResumen.getRow(1);
      headerRow1.height = 24;
      headerRow1.eachCell(c => {
        c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E293B' } };
        c.font = { name: 'Arial', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
        c.alignment = { vertical: 'middle', horizontal: 'center' };
      });

      [filaVentas, filaGastos, filaUtilidad, filaMargen, filaRecaudado, filaCartera].forEach(row => {
        row.height = 22;
        row.getCell(1).font = { name: 'Arial', size: 10, bold: true };
        row.getCell(2).font = { name: 'Arial', size: 11, bold: true };
      });

      filaVentas.getCell(2).font.color = { argb: 'FF15803D' };
      filaGastos.getCell(2).font.color = { argb: 'FFB91C1C' };
      filaUtilidad.getCell(2).font.color = { argb: 'FF1D4ED8' };

      [filaVentas, filaGastos, filaUtilidad, filaRecaudado, filaCartera].forEach(row => {
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
          inv: (d.invernaderos?.nombre || 'General').toUpperCase(),
          cliente: (d.clientes?.nombre_completo || 'Particular').toUpperCase(),
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

      // HOJA 4: DETALLE MANO DE OBRA
      const wsManoObra = workbook.addWorksheet('Detalle Mano de Obra');
      wsManoObra.columns = [
        { header: 'COMP. N°', key: 'comp', width: 14 },
        { header: 'FECHA PAGO', key: 'fecha', width: 15 },
        { header: 'INVERNADERO', key: 'inv', width: 22 },
        { header: 'COLABORADOR', key: 'nombre', width: 30 },
        { header: 'NETO PAGADO', key: 'monto', width: 20 }
      ];

      nominaFiltrada.forEach(n => {
        wsManoObra.addRow({
          comp: `NOM-${String(n.id).padStart(4, '0')}`,
          fecha: n.fecha_pago || '',
          inv: (n.invernadero_nombre || 'GENERAL / VARIOS').toUpperCase(),
          nombre: (n.nomina_trabajadores?.nombre_completo || 'OPERARIO').toUpperCase(),
          monto: parseFloat(n.monto_pagado || 0)
        });
      });

      const ultFilaM = wsManoObra.rowCount;
      const filaTotM = wsManoObra.addRow({
        nombre: 'TOTAL MANO DE OBRA:',
        monto: { formula: `=SUM(E2:E${ultFilaM})` }
      });
      filaTotM.getCell('nombre').font = { name: 'Arial', size: 10, bold: true };
      filaTotM.getCell('monto').font = { name: 'Arial', size: 11, bold: true, color: { argb: 'FF6B21A8' } };

      // FORMATEO VISUAL UNIFORME DE TABLAS
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
    <div className="space-y-6 pb-20 text-slate-800">
      
      {/* SECCIÓN DE FILTROS SUPERIORES */}
      <div className="bg-white p-6 rounded-3xl shadow-xl border-t-8 border-[#117097] space-y-4">
        <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
          
          {/* SELECTOR ORGANIZADO CON GRUPOS DE OPERATIVOS Y ARCHIVADOS */}
          <div>
            <label className="text-[10px] font-black text-gray-400 uppercase tracking-widest italic px-1">Análisis de Invernadero / Bloque</label>
            <select 
              className="w-full border-2 p-3 rounded-2xl font-black text-slate-800 bg-white mt-1 outline-none focus:border-[#117097] text-xs shadow-sm"
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

          <div className="md:col-span-2 space-y-1">
            <label className="text-[10px] font-black text-gray-400 uppercase tracking-widest italic px-1">Filtrar por Rango de Fechas</label>
            <div className="grid grid-cols-1 sm:grid-cols-2 gap-2 mt-1">
              <div className="flex items-center gap-2">
                <span className="text-[10px] font-bold text-gray-400">DESDE:</span>
                <input type="date" className="w-full border-2 p-2 rounded-xl font-bold text-xs outline-none focus:border-[#117097]" value={fechaInicio} onChange={e => { setFechaInicio(e.target.value); setFiltroPeriodoRapido('personalizado'); }} />
              </div>
              <div className="flex items-center gap-2">
                <span className="text-[10px] font-bold text-gray-400">HASTA:</span>
                <input type="date" className="w-full border-2 p-2 rounded-xl font-bold text-xs outline-none focus:border-[#117097]" value={fechaFin} onChange={e => { setFechaFin(e.target.value); setFiltroPeriodoRapido('personalizado'); }} />
              </div>
            </div>
          </div>
        </div>

        <div className="pt-2 border-t border-slate-100 flex flex-col sm:flex-row justify-between items-center gap-3">
          <div className="flex gap-1.5 flex-wrap">
            <button onClick={() => aplicarPeriodoRapido('mes_actual')} className={`px-3 py-1.5 rounded-xl text-[10px] font-black uppercase transition-all ${filtroPeriodoRapido === 'mes_actual' ? 'bg-[#117097] text-white shadow' : 'bg-slate-100 text-slate-600'}`}>📅 Mes Actual</button>
            <button onClick={() => aplicarPeriodoRapido('ano_actual')} className={`px-3 py-1.5 rounded-xl text-[10px] font-black uppercase transition-all ${filtroPeriodoRapido === 'ano_actual' ? 'bg-[#117097] text-white shadow' : 'bg-slate-100 text-slate-600'}`}>🗓️ Año 2026</button>
            <button onClick={() => aplicarPeriodoRapido('historico_total')} className={`px-3 py-1.5 rounded-xl text-[10px] font-black uppercase transition-all ${filtroPeriodoRapido === 'historico_total' ? 'bg-[#117097] text-white shadow' : 'bg-slate-100 text-slate-600'}`}>🌐 Histórico Todo</button>
          </div>

          <button
            onClick={exportarReporteAExcel}
            className="w-full sm:w-auto px-4 py-2.5 bg-emerald-600 hover:bg-emerald-700 text-white font-black rounded-xl shadow-md transition-all flex items-center justify-center gap-2 text-[10px] uppercase tracking-wider cursor-pointer"
          >
            📊 EXPORTAR A EXCEL TOTAL REPORTE DE VENTAS
          </button>
        </div>
      </div>

      {/* CARDS DINÁMICAS */}
      <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-4 gap-4">
        <div className="bg-white p-5 rounded-3xl shadow-xl border-b-4 border-red-500 text-center flex flex-col justify-center">
          <p className="text-[10px] font-black text-gray-400 uppercase tracking-wider italic">Gastos Totales (Insumos + Mano de Obra)</p>
          <p className="text-xl font-black text-red-600 mt-1">{formatoPesos(totalGastos)}</p>
        </div>
        <div className="bg-white p-5 rounded-3xl shadow-xl border-b-4 border-green-600 text-center flex flex-col justify-center">
          <p className="text-[10px] font-black text-gray-400 uppercase tracking-wider italic">Ventas Totales (Ingresos)</p>
          <p className="text-xl font-black text-green-700 mt-1">{formatoPesos(totalVentas)}</p>
        </div>
        <div className="bg-white p-5 rounded-3xl shadow-xl border-b-4 border-blue-500 text-center flex flex-col justify-center">
          <p className="text-[10px] font-black text-gray-400 uppercase tracking-wider italic">Utilidad Neta</p>
          <p className="text-xl font-black text-blue-700 mt-1">{formatoPesos(utilidadNeta)}</p>
        </div>
        <div className="bg-slate-800 p-5 rounded-3xl shadow-xl text-center flex flex-col justify-center border-b-4 border-slate-900">
          <p className="text-[10px] font-black text-slate-400 uppercase tracking-wider italic">Margen de Rendimiento</p>
          <p className="text-2xl font-black text-white mt-1">{margen}%</p>
        </div>
      </div>

      {/* GRÁFICA Y CARTERA */}
      <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
        <div className="bg-white p-6 rounded-3xl shadow-xl h-64 border border-gray-100">
          <p className="text-[10px] font-black text-slate-400 uppercase mb-4 italic tracking-wider">Comparativa Rendimiento Financiero</p>
          <ResponsiveContainer width="100%" height="100%">
            <BarChart data={dataGrafica}>
              <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f0f0f0" />
              <XAxis dataKey="name" axisLine={false} tickLine={false} tick={{fontSize: 10, fontStyle: 'bold', fill: '#475569'}} />
              <YAxis hide />
              <Tooltip formatter={(value) => formatoPesos(value)} contentStyle={{borderRadius: '15px', border: '1px solid #e2e8f0', boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1)'}} />
              <Bar dataKey="valor" radius={[8, 8, 0, 0]}>
                {dataGrafica.map((entry, index) => (
                  <Cell key={`cell-${index}`} fill={entry.color} />
                ))}
              </Bar>
            </BarChart>
          </ResponsiveContainer>
        </div>

        <div className="bg-white p-6 rounded-3xl shadow-xl flex flex-col justify-center h-64 border border-gray-100">
          <p className="text-[10px] font-black text-slate-400 uppercase mb-5 italic tracking-wider">Estado de Recaudos y Cartera</p>
          <div className="space-y-4">
            <div className="flex justify-between items-center bg-green-50 p-4 rounded-2xl border border-green-100 shadow-sm">
              <span className="text-[10px] font-black text-green-800 uppercase tracking-wider">Total Cobrado (Caja Recaudada)</span>
              <span className="font-black text-green-700 text-sm">{formatoPesos(pagosRecibidos)}</span>
            </div>
            <div className="flex justify-between items-center bg-amber-50 p-4 rounded-2xl border border-amber-100 shadow-sm">
              <span className="text-[10px] font-black text-amber-800 uppercase tracking-wider">Por Cobrar (Cartera Pendiente)</span>
              <span className="font-black text-amber-700 text-sm">{formatoPesos(cuentasPorCobrar)}</span>
            </div>
          </div>
        </div>
      </div>

      {/* VISTA EXPLORATORIA EN PANTALLA */}
      <div className="bg-white rounded-3xl shadow-xl overflow-hidden border border-gray-200">
        <div className="p-4 bg-slate-800 text-white font-black text-xs uppercase tracking-wider flex justify-between items-center flex-wrap gap-2">
          <span>🔍 Explorador de Registros Filtrados ({nombreInvSeleccionado})</span>
          <div className="flex gap-2">
            <button onClick={() => setTabDetalleAccion('ventas')} className={`px-3 py-1 rounded-lg text-[10px] uppercase font-black transition-all ${tabDetalleAccion === 'ventas' ? 'bg-[#117097] text-white' : 'bg-slate-700 text-slate-300'}`}>🛒 Ventas ({despachosFiltrados.length})</button>
            <button onClick={() => setTabDetalleAccion('gastos')} className={`px-3 py-1 rounded-lg text-[10px] uppercase font-black transition-all ${tabDetalleAccion === 'gastos' ? 'bg-[#117097] text-white' : 'bg-slate-700 text-slate-300'}`}>💸 Insumos ({egresosFiltrados.length})</button>
            <button onClick={() => setTabDetalleAccion('nomina')} className={`px-3 py-1 rounded-lg text-[10px] uppercase font-black transition-all ${tabDetalleAccion === 'nomina' ? 'bg-[#117097] text-white' : 'bg-slate-700 text-slate-300'}`}>👥 Mano Obra ({nominaFiltrada.length})</button>
          </div>
        </div>

        <div className="overflow-x-auto max-h-72 overflow-y-auto">
          {tabDetalleAccion === 'ventas' && (
            <table className="w-full text-left border-collapse text-[10px]">
              <thead>
                <tr className="bg-slate-100 text-slate-600 uppercase font-black border-b sticky top-0">
                  <th className="p-3">Fecha</th>
                  <th className="p-3">Remisión N°</th>
                  <th className="p-3">Invernadero</th>
                  <th className="p-3">Cliente</th>
                  <th className="p-3 text-right">Monto Venta</th>
                </tr>
              </thead>
              <tbody className="divide-y font-bold text-slate-700">
                {despachosFiltrados.map(d => (
                  <tr key={d.id} className="hover:bg-slate-50">
                    <td className="p-3">{d.fecha_venta}</td>
                    <td className="p-3 font-black text-[#117097]">{d.numero_remision || 'S/N'}</td>
                    <td className="p-3 uppercase">{d.invernaderos?.nombre || 'General'}</td>
                    <td className="p-3 uppercase">{d.clientes?.nombre_completo || 'Particular'}</td>
                    <td className="p-3 text-right text-emerald-700 font-black">{formatoPesos(d.total_venta)}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          )}

          {tabDetalleAccion === 'gastos' && (
            <table className="w-full text-left border-collapse text-[10px]">
              <thead>
                <tr className="bg-slate-100 text-slate-600 uppercase font-black border-b sticky top-0">
                  <th className="p-3">Fecha</th>
                  <th className="p-3">Comprobante</th>
                  <th className="p-3">Categoría</th>
                  <th className="p-3">Descripción</th>
                  <th className="p-3 text-right">Monto Gasto</th>
                </tr>
              </thead>
              <tbody className="divide-y font-bold text-slate-700">
                {egresosFiltrados.map(g => (
                  <tr key={g.id} className="hover:bg-slate-50">
                    <td className="p-3">{g.fecha_gasto}</td>
                    <td className="p-3 font-black text-slate-500">{g.numero_comprobante || 'S/N'}</td>
                    <td className="p-3 uppercase">{g.categoria}</td>
                    <td className="p-3 uppercase text-gray-500">{g.descripcion}</td>
                    <td className="p-3 text-right text-red-600 font-black">{formatoPesos(g.monto)}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          )}

          {tabDetalleAccion === 'nomina' && (
            <table className="w-full text-left border-collapse text-[10px]">
              <thead>
                <tr className="bg-slate-100 text-slate-600 uppercase font-black border-b sticky top-0">
                  <th className="p-3 text-center">Comp. N°</th>
                  <th className="p-3">Fecha Pago</th>
                  <th className="p-3">Invernadero</th>
                  <th className="p-3">Trabajador</th>
                  <th className="p-3 text-right">Neto Entregado</th>
                </tr>
              </thead>
              <tbody className="divide-y font-bold text-slate-700">
                {nominaFiltrada.map(n => (
                  <tr key={n.id} className="hover:bg-slate-50">
                    <td className="p-3 text-center font-black text-slate-500">NOM-{String(n.id).padStart(4, '0')}</td>
                    <td className="p-3">{n.fecha_pago}</td>
                    <td className="p-3 uppercase font-black text-[#117097]">{n.invernadero_nombre || 'GENERAL / VARIOS'}</td>
                    <td className="p-3 uppercase">{n.nomina_trabajadores?.nombre_completo || 'OPERARIO'}</td>
                    <td className="p-3 text-right text-purple-800 font-black">{formatoPesos(n.monto_pagado)}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          )}
        </div>
      </div>

    </div>
  );
}