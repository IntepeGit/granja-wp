import React, { useState, useEffect } from 'react';
import { BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, Legend } from 'recharts';
import { supabase } from '../lib/supabase';

export default function Dashboard({ listaInvernaderos, datosDespachos, datosEgresos, datosPagos, balancesGrafica }) {
  const [invSeleccionado, setInvSeleccionado] = useState('');
  
  const [historicoCosecha, setHistoricoCosecha] = useState([]);
  const [historicoNominaPagada, setHistoricoNominaPagada] = useState([]); 
  const [insumosCriticos, setInsumosCriticos] = useState([]);

  useEffect(() => {
    cargarDatosExtraDashboard();
  }, []);

  const cargarDatosExtraDashboard = async () => {
    try {
      const [resCosecha, resNominaPagada, resBodega] = await Promise.all([
        supabase.from('produccion_cosecha').select('*, invernaderos(nombre)'),
        supabase.from('nomina_pagos_realizados').select('*'),
        supabase.from('inventario').select('*')
      ]);

      setHistoricoCosecha(resCosecha.data || []);
      setHistoricoNominaPagada(resNominaPagada.data || []);

      if (resBodega.data && resBodega.data.length > 0) {
        const criticos = resBodega.data.filter(item => {
          const cantVal = parseFloat(
            item.cant_actual ?? 
            item.cantidad_actual ?? 
            item.cantidad ?? 
            item.stock ?? 
            0
          );

          const minVal = parseFloat(item.stock_minimo) > 0 
            ? parseFloat(item.stock_minimo) 
            : parseFloat(item.cant_minima ?? item.minimo ?? 3);

          const esConsumible = item.tipo_item === 'Consumible' || !item.tipo_item;
          const aplicaControl = item.aplica_stock !== false;

          return esConsumible && aplicaControl && (cantVal <= minVal);
        });

        setInsumosCriticos(criticos);
      }

    } catch (err) {
      console.error("Error cargando datos en Dashboard:", err);
    }
  };

  const obtenerNombreArticuloReal = (item) => {
    if (item.articulo && item.articulo.trim()) return item.articulo;
    if (item.nombre_articulo && item.nombre_articulo.trim()) return item.nombre_articulo;
    if (item.nombre_insumo && item.nombre_insumo.trim()) return item.nombre_insumo;
    if (item.nombre && item.nombre.trim()) return item.nombre;
    if (item.nombre_item && item.nombre_item.trim()) return item.nombre_item;
    if (item.item && item.item.trim()) return item.item;
    if (item.descripcion && item.descripcion.trim()) return item.descripcion;

    const claveNombre = Object.keys(item).find(key => 
      ['articulo', 'nombre', 'item', 'descripcion', 'producto'].some(k => key.toLowerCase().includes(k))
    );
    
    if (claveNombre && item[claveNombre]) return String(item[claveNombre]);
    return 'INSUMO';
  };

  const formatoPesos = (valor) => new Intl.NumberFormat('es-CO', { 
    style: 'currency', currency: 'COP', minimumFractionDigits: 0 
  }).format(valor || 0);

  // --- 1. SEPARACIÓN ESTRICTA DE INVERNADEROS OPERATIVOS ---
  const invernaderosOperativos = (listaInvernaderos || []).filter(inv => inv.activo !== false);
  const idsOperativos = invernaderosOperativos.map(i => i.id?.toString());
  const nombresOperativos = invernaderosOperativos.map(i => i.nombre?.toUpperCase());
  const cantidadInvernaderosActivos = invernaderosOperativos.length || 1;

  // --- 🧮 2. CÁLCULO DE NÓMINA PAGADA ---
  const calcularNominaPorInvernadero = (invId) => {
    const invObj = invernaderosOperativos.find(i => i.id?.toString() === invId?.toString());
    const nombreInv = invObj?.nombre?.toUpperCase();

    let totalManoObraLote = 0;

    (historicoNominaPagada || []).forEach(pago => {
      const monto = parseFloat(pago.monto_pagado || 0);
      const invNombrePago = (pago.invernadero_nombre || '').toUpperCase();

      if (invNombrePago && nombreInv && invNombrePago.includes(nombreInv)) {
        totalManoObraLote += monto;
      } 
      else if (!invNombrePago || invNombrePago.includes('GENERAL') || invNombrePago.includes('VARIOS')) {
        totalManoObraLote += (monto / cantidadInvernaderosActivos);
      }
    });

    return totalManoObraLote;
  };

  const gastosNominaGlobales = (historicoNominaPagada || []).reduce((acc, p) => acc + parseFloat(p.monto_pagado || 0), 0);

  // --- 🧮 3. SUMATORIAS GLOBALES ---
  const ingresosGlobales = (datosDespachos || [])
    .filter(d => idsOperativos.includes(d.invernadero_id?.toString()))
    .reduce((acc, d) => acc + (parseFloat(d.total_venta) || 0), 0);

  const gastosInsumosGlobales = (datosEgresos || [])
    .filter(g => !g.invernadero_id || idsOperativos.includes(g.invernadero_id?.toString()))
    .reduce((acc, e) => acc + (parseFloat(e.monto) || 0), 0);

  const gastosTotalesConNomina = gastosInsumosGlobales + gastosNominaGlobales;
  const utilidadRealGlobal = ingresosGlobales - gastosTotalesConNomina;

  // --- 📊 4. MATEMÁTICA DE PUNTO DE EQUILIBRIO POR INVERNADERO ---
  const analisisPuntoEquilibrio = invernaderosOperativos.map(inv => {
    const gastosInsumosLote = (datosEgresos || [])
      .filter(g => g.invernadero_id?.toString() === inv.id?.toString())
      .reduce((acc, g) => acc + (parseFloat(g.monto) || 0), 0);

    const gastosNominaLote = calcularNominaPorInvernadero(inv.id);
    const inversionTotalLote = gastosInsumosLote + gastosNominaLote;
    
    const cosechasLote = (historicoCosecha || []).filter(c => c.invernadero_id === inv.id);
    
    const totalCanastillas = cosechasLote
      .filter(c => c.unidad_medida === 'CANASTILLA')
      .reduce((acc, c) => acc + (parseFloat(c.cantidad) || 0), 0);

    const otrosEmpaques = cosechasLote
      .filter(c => c.unidad_medida !== 'CANASTILLA')
      .reduce((acc, c) => acc + (parseFloat(c.cantidad) || 0), 0);

    const totalUnidadesCosechadas = totalCanastillas > 0 ? totalCanastillas : otrosEmpaques;
    const etiquetaUnidad = totalCanastillas > 0 ? 'Canastillas' : 'Unidades / Kilos';

    const costoBaseUnidad = totalUnidadesCosechadas > 0 ? (inversionTotalLote / totalUnidadesCosechadas) : 0;

    return {
      id: inv.id,
      nombre: inv.nombre?.toUpperCase(),
      inversionTotal: inversionTotalLote,
      totalUnidades: totalUnidadesCosechadas,
      etiquetaUnidad,
      costoBaseUnidad
    };
  });

  // --- 5. DATOS PARA GRÁFICOS ---
  const balancesGraficaOperativos = (balancesGrafica || []).filter(b => nombresOperativos.includes(b.name?.toUpperCase()));

  const datosGraficoProduccion = invernaderosOperativos.map(inv => {
    const cosechasLote = historicoCosecha.filter(c => c.invernadero_id === inv.id);
    const canastillas = cosechasLote.filter(c => c.unidad_medida === 'CANASTILLA').reduce((acc, c) => acc + (parseFloat(c.cantidad) || 0), 0);
    const bultos = cosechasLote.filter(c => c.unidad_medida === 'BULTO').reduce((acc, c) => acc + (parseFloat(c.cantidad) || 0), 0);
    return { name: inv.nombre?.toUpperCase(), Canastillas: canastillas, Bultos: bultos };
  });

  const datosGraficoNomina = invernaderosOperativos.map(inv => {
    const costoNominaLote = calcularNominaPorInvernadero(inv.id);
    return { name: inv.nombre?.toUpperCase(), Costo_Mano_Obra: costoNominaLote };
  });

  // --- 6. FILTRADO EXPLORADOR ---
  const despachosInv = datosDespachos?.filter(d => d.invernadero_id?.toString() === invSeleccionado) || [];
  const gastosInv = datosEgresos?.filter(g => g.invernadero_id?.toString() === invSeleccionado) || [];
  
  const objInvSel = invernaderosOperativos.find(i => i.id?.toString() === invSeleccionado);
  const nombreInvSelStr = objInvSel ? objInvSel.nombre?.toUpperCase() : '';

  const nominaInv = historicoNominaPagada?.filter(p => {
    const invPagoStr = (p.invernadero_nombre || '').toUpperCase();
    return invPagoStr.includes(nombreInvSelStr) || invPagoStr.includes('GENERAL') || invPagoStr.includes('VARIOS');
  }) || [];
  
  const idsDespachos = despachosInv.map(d => d.id?.toString());
  const pagosInv = datosPagos?.filter(p => idsDespachos.includes(p.despacho_id?.toString())) || [];

  const totalRemisiones = despachosInv.reduce((acc, d) => acc + (d.total_venta || 0), 0);
  const totalAbonos = pagosInv.reduce((acc, p) => acc + (p.monto || 0), 0);
  const totalGastosInsumos = gastosInv.reduce((acc, g) => acc + (g.monto || 0), 0);
  const totalManoObra = calcularNominaPorInvernadero(invSeleccionado);

  return (
    <div className="space-y-4 animate-in fade-in duration-500 pb-10 text-slate-800 dark:text-slate-200 font-sans transition-colors duration-300">
      
      {/* 🚀 CABECERA PRINCIPAL MODERNA */}
      <div className="bg-white dark:bg-slate-800 p-6 rounded-3xl shadow-xl border border-slate-200 dark:border-slate-700 flex flex-col md:flex-row justify-between items-center gap-4 transition-colors duration-300">
        <div className="space-y-1 w-full">
          <div className="flex items-center gap-3">
            <span className="text-2xl p-2 bg-[#117097]/10 dark:bg-sky-500/20 rounded-xl text-[#117097] dark:text-sky-400">📊</span>
            <h2 className="text-xl font-black tracking-tight text-slate-900 dark:text-white uppercase">Dashboard General</h2>
          </div>
          <p className="text-xs text-slate-500 dark:text-slate-400 font-medium">Análisis en tiempo real de operaciones, finanzas y estado de la granja.</p>
        </div>
      </div>

      {/* 🚨 BANNER DE ALERTA DE STOCK */}
      {insumosCriticos.length > 0 && (
        <div className="bg-amber-50 dark:bg-amber-950/40 border border-amber-300 dark:border-amber-800/60 p-4 rounded-2xl shadow-sm flex flex-col sm:flex-row items-start sm:items-center justify-between gap-4 transition-colors duration-300">
          <div className="flex items-center gap-3">
            <span className="text-2xl animate-pulse">⚠️</span>
            <div>
              <h4 className="font-black text-amber-900 dark:text-amber-500 text-xs uppercase tracking-wider">
                Alerta de Reabastecimiento en Bodega ({insumosCriticos.length} Críticos)
              </h4>
              <div className="flex gap-2 flex-wrap mt-1.5">
                {insumosCriticos.map((item, idx) => {
                  const nombreArticulo = obtenerNombreArticuloReal(item);
                  const cantidadStock = item.cantidad_actual ?? item.cant_actual ?? item.cantidad ?? item.stock ?? 0;
                  const minDef = parseFloat(item.stock_minimo) > 0 ? parseFloat(item.stock_minimo) : parseFloat(item.cant_minima ?? item.minimo ?? 3);
                  
                  return (
                    <span key={item.id || idx} className="bg-amber-200/50 dark:bg-amber-900/50 text-amber-900 dark:text-amber-200 border border-amber-300 dark:border-amber-700 px-2 py-1 rounded-md font-black text-[9px] uppercase shadow-sm">
                      📦 {nombreArticulo}: <span className="text-red-700 dark:text-red-400 font-extrabold">{cantidadStock}</span> (Mín: {minDef})
                    </span>
                  );
                })}
              </div>
            </div>
          </div>
        </div>
      )}

      {/* 📈 KPIs GLOBALES */}
      <div className="grid grid-cols-2 lg:grid-cols-4 gap-4">
        <div className="bg-white dark:bg-slate-800 p-5 rounded-2xl shadow-sm border border-slate-100 dark:border-slate-700 border-l-4 border-l-[#117097] dark:border-l-sky-500 transition-colors duration-300">
          <p className="text-[9px] font-black text-slate-400 dark:text-slate-400 uppercase tracking-wider">Ingresos Totales (En Producción)</p>
          <p className="text-2xl font-black text-slate-800 dark:text-white mt-1">{formatoPesos(ingresosGlobales)}</p>
        </div>
        <div className="bg-white dark:bg-slate-800 p-5 rounded-2xl shadow-sm border border-slate-100 dark:border-slate-700 border-l-4 border-l-slate-400 dark:border-l-slate-500 transition-colors duration-300">
          <p className="text-[9px] font-black text-slate-400 dark:text-slate-400 uppercase tracking-wider">Gastos (+Mano Obra Liquidada)</p>
          <p className="text-2xl font-black text-slate-800 dark:text-white mt-1">{formatoPesos(gastosTotalesConNomina)}</p>
        </div>
        <div className="bg-white dark:bg-slate-800 p-5 rounded-2xl shadow-sm border border-slate-100 dark:border-slate-700 border-l-4 border-l-emerald-500 transition-colors duration-300">
          <p className="text-[9px] font-black text-slate-400 dark:text-slate-400 uppercase tracking-wider">Utilidad Neta Activa</p>
          <p className={`text-2xl font-black mt-1 ${utilidadRealGlobal >= 0 ? 'text-emerald-600 dark:text-emerald-400' : 'text-rose-600 dark:text-rose-500'}`}>
            {formatoPesos(utilidadRealGlobal)}
          </p>
        </div>
        <div className="bg-[#117097] dark:bg-sky-900/50 p-5 rounded-2xl shadow-md text-white border-l-4 border-l-sky-300 dark:border-l-sky-400 transition-colors duration-300">
          <p className="text-[9px] font-black text-sky-100 dark:text-sky-200 uppercase tracking-wider">Eficiencia Financiera</p>
          <p className="text-2xl font-black mt-1">
            {ingresosGlobales > 0 ? ((utilidadRealGlobal / ingresosGlobales) * 100).toFixed(1) : 0}%
          </p>
        </div>
      </div>

      {/* 🎯 COSTO BASE Y PUNTO DE EQUILIBRIO */}
      <div className="bg-white dark:bg-slate-800 p-5 rounded-3xl shadow-sm border border-slate-200 dark:border-slate-700 transition-colors duration-300">
        <div className="flex justify-between items-center mb-4 border-b border-slate-100 dark:border-slate-700 pb-2">
          <h3 className="font-black text-slate-700 dark:text-slate-200 text-xs uppercase tracking-wider flex items-center gap-1.5">
            🎯 Costo Base y Punto de Equilibrio por Lote
          </h3>
          <span className="text-[9px] font-bold text-slate-500 dark:text-slate-400 uppercase bg-slate-100 dark:bg-slate-700 px-2 py-0.5 rounded">
            Lotes Activos
          </span>
        </div>

        <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
          {analisisPuntoEquilibrio.map((item) => (
            <div key={item.id} className="p-4 bg-slate-50 dark:bg-slate-800 rounded-2xl border border-slate-100 dark:border-slate-700 flex flex-col justify-between hover:border-[#117097]/30 dark:hover:border-sky-500/50 transition-colors">
              <div>
                <div className="flex justify-between items-center mb-2 border-b border-dashed border-slate-200 dark:border-slate-700 pb-2">
                  <p className="font-black text-sm text-[#117097] dark:text-sky-400 uppercase">{item.nombre}</p>
                  <span className="text-[9px] font-bold text-slate-400 dark:text-slate-500 uppercase">{item.etiquetaUnidad}</span>
                </div>

                <div className="space-y-1.5 text-[10px] font-bold text-slate-600 dark:text-slate-400">
                  <div className="flex justify-between">
                    <span>Inversión Lote:</span>
                    <span className="font-black text-slate-800 dark:text-slate-200">{formatoPesos(item.inversionTotal)}</span>
                  </div>
                  <div className="flex justify-between">
                    <span>Cosechado:</span>
                    <span className="font-black text-emerald-700 dark:text-emerald-400">{item.totalUnidades} {item.etiquetaUnidad}</span>
                  </div>
                </div>
              </div>

              <div className="bg-amber-50/50 dark:bg-amber-900/20 p-3 rounded-xl border border-amber-100 dark:border-amber-800/50 mt-3 text-center">
                <p className="text-[9px] font-black text-amber-700 dark:text-amber-500 uppercase tracking-wider">Precio Mínimo de Venta Base</p>
                <p className="text-base font-black text-amber-900 dark:text-amber-400 mt-1">
                  {item.costoBaseUnidad > 0 ? `${formatoPesos(item.costoBaseUnidad)} / U` : 'Sin Cosecha'}
                </p>
              </div>
            </div>
          ))}
        </div>
      </div>

      {/* 📊 GRÁFICO PRINCIPAL COMERCIAL + BALANCES INDIVIDUALES */}
      <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
        <div className="lg:col-span-2 bg-white dark:bg-slate-800 p-5 rounded-3xl shadow-sm border border-slate-200 dark:border-slate-700 transition-colors duration-300">
          <h3 className="font-black text-slate-700 dark:text-slate-200 text-xs uppercase tracking-wider mb-4 border-b border-slate-100 dark:border-slate-700 pb-2">Balance Financiero por Invernadero (En Producción)</h3>
          <div className="h-64 w-full">
            <ResponsiveContainer width="100%" height="100%">
              <BarChart data={balancesGraficaOperativos} margin={{ top: 5, right: 5, left: -20, bottom: 0 }}>
                <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#334155" opacity={0.2} />
                <XAxis dataKey="name" fontSize={10} fontWeight="800" axisLine={false} tickLine={false} tick={{fill: '#94a3b8'}} />
                <YAxis fontSize={10} axisLine={false} tickLine={false} tickFormatter={(v) => `$${v/1000000}M`} tick={{fill: '#94a3b8'}} />
                <Tooltip contentStyle={{ borderRadius: '12px', border: '1px solid #334155', backgroundColor: '#1e293b', color: '#fff', fontSize: '11px', fontWeight: 'bold' }} formatter={(v) => formatoPesos(v)} />
                <Legend verticalAlign="top" height={24} iconType="circle" wrapperStyle={{ fontSize: '10px', fontWeight: 'bold', color: '#94a3b8' }} />
                <Bar dataKey="Ingresos" fill="#117097" radius={[4, 4, 0, 0]} name="Ventas" />
                <Bar dataKey="Gastos" fill="#94a3b8" radius={[4, 4, 0, 0]} name="Gastos/Insumos" />
                <Bar dataKey="Utilidad" fill="#059669" radius={[4, 4, 0, 0]} name="Margen" />
              </BarChart>
            </ResponsiveContainer>
          </div>
        </div>

        <div className="bg-white dark:bg-slate-800 p-5 rounded-3xl shadow-sm border border-slate-200 dark:border-slate-700 flex flex-col max-h-[350px] transition-colors duration-300">
          <h3 className="font-black text-slate-700 dark:text-slate-200 text-xs uppercase tracking-wider mb-4 border-b border-slate-100 dark:border-slate-700 pb-2">Balances Individuales</h3>
          <div className="space-y-2 flex-1 overflow-y-auto pr-2 custom-scrollbar">
            {balancesGraficaOperativos.map((inv, idx) => {
              const objetoInvReal = invernaderosOperativos.find(i => i.nombre?.toUpperCase() === inv.name?.toUpperCase());
              const nominaEsteLote = objetoInvReal ? calcularNominaPorInvernadero(objetoInvReal.id) : 0;
              const gastosTotalesLote = inv.Gastos + nominaEsteLote;
              const utilidadRealLote = inv.Ingresos - gastosTotalesLote;

              return (
                <div key={idx} className="p-3 bg-slate-50 dark:bg-slate-700/50 rounded-xl border border-slate-100 dark:border-slate-700 flex items-center justify-between hover:bg-slate-100 dark:hover:bg-slate-700 transition-colors">
                  <div>
                    <p className="font-black text-sm text-[#117097] dark:text-sky-400 uppercase tracking-tight">{inv.name}</p>
                    <p className="text-[10px] text-gray-500 dark:text-gray-400 font-bold mt-0.5">Cartera: {formatoPesos(inv.Cartera)}</p>
                  </div>
                  <div className="text-right">
                    <p className="font-black text-slate-800 dark:text-white text-sm">{formatoPesos(inv.Ingresos)}</p>
                    <p className={`text-[10px] font-black uppercase mt-0.5 ${utilidadRealLote >= 0 ? 'text-emerald-600 dark:text-emerald-400' : 'text-rose-600 dark:text-rose-400'}`}>
                      Util: {formatoPesos(utilidadRealLote)}
                    </p>
                  </div>
                </div>
              );
            })}
          </div>
        </div>
      </div>

      {/* 🚜 GRÁFICOS SECUNDARIOS */}
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
        <div className="bg-white dark:bg-slate-800 p-5 rounded-3xl shadow-sm border border-slate-200 dark:border-slate-700 transition-colors duration-300">
          <h4 className="font-black text-slate-700 dark:text-slate-200 text-xs uppercase tracking-wider mb-4 border-b border-slate-100 dark:border-slate-700 pb-2">🚜 Volumen Cosechado por Lote</h4>
          <div className="h-56 w-full">
            <ResponsiveContainer width="100%" height="100%">
              <BarChart data={datosGraficoProduccion} margin={{ top: 5, right: 5, left: -25, bottom: 0 }}>
                <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#334155" opacity={0.2} />
                <XAxis dataKey="name" fontSize={10} fontWeight="800" axisLine={false} tickLine={false} tick={{fill: '#94a3b8'}} />
                <YAxis fontSize={10} axisLine={false} tickLine={false} tick={{fill: '#94a3b8'}} />
                <Tooltip contentStyle={{ borderRadius: '12px', border: '1px solid #334155', backgroundColor: '#1e293b', color: '#fff', fontSize: '11px', fontWeight: 'bold' }} />
                <Legend iconType="square" wrapperStyle={{ fontSize: '10px', fontWeight: 'bold', color: '#94a3b8' }} />
                <Bar dataKey="Canastillas" fill="#059669" radius={[4, 4, 0, 0]} />
                <Bar dataKey="Bultos" fill="#b45309" radius={[4, 4, 0, 0]} />
              </BarChart>
            </ResponsiveContainer>
          </div>
        </div>

        <div className="bg-white dark:bg-slate-800 p-5 rounded-3xl shadow-sm border border-slate-200 dark:border-slate-700 transition-colors duration-300">
          <h4 className="font-black text-slate-700 dark:text-slate-200 text-xs uppercase tracking-wider mb-4 border-b border-slate-100 dark:border-slate-700 pb-2">👥 Costo de Mano de Obra por Lote</h4>
          <div className="h-56 w-full">
            <ResponsiveContainer width="100%" height="100%">
              <BarChart data={datosGraficoNomina} margin={{ top: 5, right: 5, left: -20, bottom: 0 }}>
                <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#334155" opacity={0.2} />
                <XAxis dataKey="name" fontSize={10} fontWeight="800" axisLine={false} tickLine={false} tick={{fill: '#94a3b8'}} />
                <YAxis fontSize={10} axisLine={false} tickLine={false} tickFormatter={(v) => `$${v/1000}K`} tick={{fill: '#94a3b8'}} />
                <Tooltip contentStyle={{ borderRadius: '12px', border: '1px solid #334155', backgroundColor: '#1e293b', color: '#fff', fontSize: '11px', fontWeight: 'bold' }} formatter={(v) => formatoPesos(v)} />
                <Bar dataKey="Costo_Mano_Obra" fill="#0284c7" radius={[4, 4, 0, 0]} name="Costo Jornales / Sueldos" />
              </BarChart>
            </ResponsiveContainer>
          </div>
        </div>
      </div>

      {/* 🔍 EXPLORADOR DE DETALLES */}
      <div className="bg-slate-50 dark:bg-slate-800/50 p-5 rounded-3xl border border-slate-200 dark:border-slate-700 transition-colors duration-300">
        <div className="flex flex-col sm:flex-row justify-between items-center gap-4 mb-5 border-b border-slate-200 dark:border-slate-700 pb-4">
          <h3 className="font-black text-slate-700 dark:text-slate-200 text-xs uppercase tracking-wider flex items-center gap-2">
            <span>🔍</span> Explorador de Detalles
          </h3>
          <select 
            value={invSeleccionado} 
            onChange={(e) => setInvSeleccionado(e.target.value)}
            className="w-full sm:w-auto p-2.5 rounded-xl border border-slate-300 dark:border-slate-600 font-bold text-xs outline-none focus:border-[#117097] dark:focus:border-sky-500 bg-white dark:bg-slate-700 dark:text-white transition-colors cursor-pointer shadow-sm"
          >
            <option value="">Seleccione Invernadero para auditar...</option>
            <optgroup label="🌱 EN PRODUCCIÓN (OPERATIVOS)">
              {(listaInvernaderos || []).filter(i => i.activo !== false).map(i => (
                <option key={i.id} value={i.id}>{i.nombre?.toUpperCase()}</option>
              ))}
            </optgroup>
            <optgroup label="📁 HISTÓRICO / ARCHIVADOS">
              {(listaInvernaderos || []).filter(i => i.activo === false).map(i => (
                <option key={i.id} value={i.id}>{i.nombre?.toUpperCase()} (ARCHIVADO)</option>
              ))}
            </optgroup>
          </select>
        </div>

        {invSeleccionado ? (
          <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
            
            <div className="bg-white dark:bg-slate-800 rounded-2xl shadow-sm border border-slate-200 dark:border-slate-700 overflow-hidden transition-colors duration-300">
              <div className="bg-[#117097] p-3 text-white flex justify-between items-center">
                <span className="font-black text-[10px] uppercase tracking-wider">🛒 Remisiones</span>
                <span className="bg-white/20 px-2.5 py-1 rounded text-[10px] font-black">{formatoPesos(totalRemisiones)}</span>
              </div>
              <div className="max-h-48 overflow-y-auto p-2">
                <table className="w-full text-[10px]">
                  <tbody className="divide-y divide-slate-100 dark:divide-slate-700">
                    {despachosInv.map(d => (
                      <tr key={d.id} className="hover:bg-slate-50 dark:hover:bg-slate-700/50 transition-colors">
                        <td className="p-2.5 font-bold uppercase text-slate-700 dark:text-slate-300">#{d.numero_remision}</td>
                        <td className="p-2.5 text-right font-black text-[#117097] dark:text-sky-400">{formatoPesos(d.total_venta)}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
                {despachosInv.length === 0 && <p className="text-center text-[10px] font-bold text-slate-400 dark:text-slate-500 py-4">Sin remisiones</p>}
              </div>
            </div>

            <div className="bg-white dark:bg-slate-800 rounded-2xl shadow-sm border border-slate-200 dark:border-slate-700 overflow-hidden transition-colors duration-300">
              <div className="bg-amber-500 dark:bg-amber-600 p-3 text-white flex justify-between items-center">
                <span className="font-black text-[10px] uppercase tracking-wider">💰 Abonos</span>
                <span className="bg-white/20 px-2.5 py-1 rounded text-[10px] font-black">{formatoPesos(totalAbonos)}</span>
              </div>
              <div className="max-h-48 overflow-y-auto p-2">
                <table className="w-full text-[10px]">
                  <tbody className="divide-y divide-slate-100 dark:divide-slate-700">
                    {pagosInv.map(p => (
                      <tr key={p.id} className="hover:bg-slate-50 dark:hover:bg-slate-700/50 transition-colors">
                        <td className="p-2.5 font-bold uppercase text-slate-700 dark:text-slate-300">#{p.ventas?.numero_remision}</td>
                        <td className="p-2.5 text-right font-black text-amber-600 dark:text-amber-400">{formatoPesos(p.monto)}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
                {pagosInv.length === 0 && <p className="text-center text-[10px] font-bold text-slate-400 dark:text-slate-500 py-4">Sin abonos</p>}
              </div>
            </div>

            <div className="bg-white dark:bg-slate-800 rounded-2xl shadow-sm border border-slate-200 dark:border-slate-700 overflow-hidden transition-colors duration-300">
              <div className="bg-slate-500 dark:bg-slate-600 p-3 text-white flex justify-between items-center">
                <span className="font-black text-[10px] uppercase tracking-wider">🧪 Gastos Insumos</span>
                <span className="bg-white/20 px-2.5 py-1 rounded text-[10px] font-black">{formatoPesos(totalGastosInsumos)}</span>
              </div>
              <div className="max-h-48 overflow-y-auto p-2">
                <table className="w-full text-[10px]">
                  <tbody className="divide-y divide-slate-100 dark:divide-slate-700">
                    {gastosInv.map(g => (
                      <tr key={g.id} className="hover:bg-slate-50 dark:hover:bg-slate-700/50 transition-colors">
                        <td className="p-2.5 font-bold uppercase text-slate-700 dark:text-slate-300 truncate max-w-[80px]" title={g.descripcion}>{g.descripcion}</td>
                        <td className="p-2.5 text-right font-black text-slate-600 dark:text-slate-400">{formatoPesos(g.monto)}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
                {gastosInv.length === 0 && <p className="text-center text-[10px] font-bold text-slate-400 dark:text-slate-500 py-4">Sin gastos</p>}
              </div>
            </div>

            <div className="bg-white dark:bg-slate-800 rounded-2xl shadow-sm border border-slate-200 dark:border-slate-700 overflow-hidden transition-colors duration-300">
              <div className="bg-blue-600 dark:bg-blue-700 p-3 text-white flex justify-between items-center">
                <span className="font-black text-[10px] uppercase tracking-wider">👥 Mano de Obra</span>
                <span className="bg-white/20 px-2.5 py-1 rounded text-[10px] font-black">{formatoPesos(totalManoObra)}</span>
              </div>
              <div className="max-h-48 overflow-y-auto p-2">
                <table className="w-full text-[10px]">
                  <tbody className="divide-y divide-slate-100 dark:divide-slate-700">
                    {nominaInv.map(n => (
                      <tr key={n.id} className="hover:bg-slate-50 dark:hover:bg-slate-700/50 transition-colors">
                        <td className="p-2.5 font-bold uppercase text-slate-700 dark:text-slate-300 truncate max-w-[80px]" title={n.invernadero_nombre}>{n.invernadero_nombre || 'GENERAL'}</td>
                        <td className="p-2.5 text-right font-black text-blue-700 dark:text-blue-400">{formatoPesos(n.monto_pagado)}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
                {nominaInv.length === 0 && <p className="text-center text-[10px] font-bold text-slate-400 dark:text-slate-500 py-4">Sin nómina asignada</p>}
              </div>
            </div>

          </div>
        ) : (
          <div className="text-center py-10 bg-white dark:bg-slate-800 rounded-2xl border border-dashed border-slate-300 dark:border-slate-600 transition-colors duration-300">
            <span className="text-4xl grayscale opacity-30 dark:opacity-20 block mb-3">🔎</span>
            <p className="text-slate-400 dark:text-slate-500 font-black uppercase text-xs tracking-widest">Seleccione un lote para auditar detalles</p>
          </div>
        )}
      </div>

      {/* 💰 BALANCE NETO REAL (PIE DE PÁGINA) */}
      <div className={`p-6 rounded-3xl shadow-xl border flex flex-col md:flex-row items-center justify-between gap-4 transition-all duration-300 ${utilidadRealGlobal >= 0 ? 'bg-[#0f4c68] dark:bg-slate-800 border-[#117097] dark:border-sky-700 text-white' : 'bg-rose-950 dark:bg-slate-800 border-rose-800 dark:border-rose-700 text-white'}`}>
        <div className="flex items-center gap-4 text-center md:text-left">
          <div className="bg-white/10 dark:bg-white/5 p-3 rounded-2xl text-3xl backdrop-blur-md shadow-inner hidden md:block">
            {utilidadRealGlobal >= 0 ? '💰' : '⚠️'}
          </div>
          <div>
            <h4 className="font-black uppercase text-sm tracking-wider">Balance de Operación Neto Real (En Producción)</h4>
            <p className="text-[10px] font-bold text-white/60 dark:text-slate-400 uppercase mt-1">Ingresos deduciendo Insumos y Mano de Obra Liquidada</p>
          </div>
        </div>
        <div className="text-center md:text-right w-full md:w-auto bg-black/20 dark:bg-black/40 p-4 rounded-2xl border border-white/10 dark:border-white/5 shadow-inner">
          <p className={`text-4xl font-black tracking-tight drop-shadow-md ${utilidadRealGlobal < 0 && 'text-rose-400'}`}>
            {formatoPesos(utilidadRealGlobal)}
          </p>
        </div>
      </div>

    </div>
  );
}