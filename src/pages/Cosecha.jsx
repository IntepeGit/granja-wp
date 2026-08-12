import React, { useState, useEffect } from 'react';
import { supabase } from '../lib/supabase';
import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';

export default function Cosecha({ mostrarAlerta, listaInvernaderos, userRole }) {
  // Estados para datos de la tabla e historial
  const [registrosCosecha, setRegistrosCosecha] = useState([]);
  const [listaTrabajadores, setListaTrabajadores] = useState([]);
  const [cargando, setCargando] = useState(false);
  const [busqueda, setBusqueda] = useState('');
  const [modalAbierto, setModalAbierto] = useState(false); 
  
  // FILTROS DE CONTROL SUPERIOR
  const [filtroInvernadero, setFiltroInvernadero] = useState('');
  const [filtroProducto, setFiltroProducto] = useState('');

  // Listas dinámicas traídas de Supabase
  const [listaProductos, setListaProductos] = useState([]);
  const [listaCalidades, setListaCalidades] = useState([]);
  const [listaUnidades, setListaUnidades] = useState([]);

  // Estado para el formulario de registro diario
  const [formCosecha, setFormCosecha] = useState({
    fecha_cosecha: new Date().toISOString().split('T')[0],
    invernadero_id: '',
    producto: '', 
    calidad: '',
    cantidad: '',
    unidad_medida: '',
    operario_recolector: '',
    observaciones: ''
  });

  useEffect(() => {
    cargarDatosIniciales();
  }, []);

  // --- 1. CARGAR CONFIGURACIONES, REGISTROS Y OPERARIOS EN PARALELO ---
  const cargarDatosIniciales = async () => {
    setCargando(true);
    try {
      const [resCosechas, resProd, resCal, resUni, resTrab] = await Promise.all([
        supabase.from('produccion_cosecha').select('*, invernaderos(nombre)').order('fecha_cosecha', { ascending: false }),
        supabase.from('config_productos').select('*').order('nombre_producto', { ascending: true }),
        supabase.from('config_calidades').select('*').order('nombre_calidad', { ascending: true }),
        supabase.from('config_unidades').select('*').order('nombre_unidad', { ascending: true }),
        supabase.from('nomina_trabajadores').select('id, nombre_completo').eq('activo', true).order('nombre_completo', { ascending: true })
      ]);

      if (resCosechas.error) throw resCosechas.error;
      if (resProd.error) throw resProd.error;
      if (resCal.error) throw resCal.error;
      if (resUni.error) throw resUni.error;

      const prods = resProd.data || [];
      const cals = resCal.data || [];
      const unis = resUni.data || [];
      const trabs = resTrab.data || [];

      setListaProductos(prods);
      setListaCalidades(cals);
      setListaUnidades(unis);
      setListaTrabajadores(trabs);
      setRegistrosCosecha(resCosechas.data || []);

      if (prods.length > 0) setFiltroProducto(prods[0].nombre_producto);

      setFormCosecha(prev => ({
        ...prev,
        producto: prev.producto || prods[0]?.nombre_producto || '',
        calidad: prev.calidad || cals[0]?.nombre_calidad || '',
        unidad_medida: prev.unidad_medida || unis[0]?.nombre_unidad || ''
      }));

    } catch (err) {
      console.error("Error cargando datos de cosecha:", err);
      if (mostrarAlerta) mostrarAlerta("No se pudieron sincronizar los parámetros", "error");
    } finally {
      setCargando(false);
    }
  };

  const abrirModalNuevo = () => {
    setFormCosecha(prev => ({ 
      ...prev, 
      fecha_cosecha: new Date().toISOString().split('T')[0],
      cantidad: '',
      observaciones: '' 
    }));
    setModalAbierto(true);
  };

  // --- 2. GUARDAR REGISTRO DIARIO DE COSECHA ---
  const guardarRegistroCosecha = async (e) => {
    e.preventDefault();
    const cantNum = parseFloat(formCosecha.cantidad);
    if (!formCosecha.invernadero_id || cantNum <= 0) {
      mostrarAlerta("Por favor completa los campos obligatorios", "error");
      return;
    }

    try {
      const { error } = await supabase
        .from('produccion_cosecha')
        .insert([{
          fecha_cosecha: formCosecha.fecha_cosecha,
          invernadero_id: formCosecha.invernadero_id,
          producto: formCosecha.producto,
          calidad: formCosecha.calidad,
          cantidad: cantNum,
          unidad_medida: formCosecha.unidad_medida,
          operario_recolector: (formCosecha.operario_recolector || 'GENERAL').toUpperCase(),
          observaciones: (formCosecha.observaciones || '').toUpperCase()
        }]);

      if (error) throw error;

      mostrarAlerta("Cosecha registrada exitosamente", "exito");
      setModalAbierto(false); 
      setFormCosecha(prev => ({ ...prev, cantidad: '', observaciones: '' }));
      
      const { data } = await supabase.from('produccion_cosecha').select('*, invernaderos(nombre)').order('fecha_cosecha', { ascending: false });
      setRegistrosCosecha(data || []);
    } catch (err) {
      mostrarAlerta("No se pudo guardar la cosecha", "error");
    }
  };

  // --- 3. ELIMINAR REGISTRO ---
  const eliminarRegistroCosecha = async (id) => {
    if (!window.confirm("¿Seguro que deseas eliminar este registro?")) return;
    try {
      const { error } = await supabase.from('produccion_cosecha').delete().eq('id', id);
      if (error) throw error;
      mostrarAlerta("Registro eliminado", "exito");
      const { data } = await supabase.from('produccion_cosecha').select('*, invernaderos(nombre)').order('fecha_cosecha', { ascending: false });
      setRegistrosCosecha(data || []);
    } catch (err) {
      mostrarAlerta("Error al intentar eliminar", "error");
    }
  };

  // --- 📊 4. EXPORTAR REGISTROS DE COSECHA A EXCEL ---
  const exportarCosechaAExcel = async () => {
    if (!registrosFiltrados || registrosFiltrados.length === 0) {
      mostrarAlerta("No hay registros de cosecha para exportar", "error");
      return;
    }

    try {
      const workbook = new ExcelJS.Workbook();
      const sheet = workbook.addWorksheet('Reporte Cosecha');

      sheet.columns = [
        { header: 'FECHA COSECHA', key: 'fecha', width: 16 },
        { header: 'INVERNADERO / BLOQUE', key: 'invernadero', width: 22 },
        { header: 'PRODUCTO', key: 'producto', width: 22 },
        { header: 'CALIDAD', key: 'calidad', width: 16 },
        { header: 'CANTIDAD RECOGIDA', key: 'cantidad', width: 20 },
        { header: 'UNIDAD MEDIDA', key: 'unidad', width: 16 },
        { header: 'OPERARIO RECOLECTOR', key: 'operario', width: 28 },
        { header: 'OBSERVACIONES', key: 'obs', width: 30 }
      ];

      registrosFiltrados.forEach(r => {
        sheet.addRow({
          fecha: r.fecha_cosecha || '',
          invernadero: (r.invernaderos?.nombre || 'GENERAL').toUpperCase(),
          producto: (r.producto || '').toUpperCase(),
          calidad: (r.calidad || '').toUpperCase(),
          cantidad: parseFloat(r.cantidad || 0),
          unidad: (r.unidad_medida || 'UNIDAD').toUpperCase(),
          operario: (r.operario_recolector || 'N/R').toUpperCase(),
          obs: (r.observaciones || 'S/N').toUpperCase()
        });
      });

      const ultFila = sheet.rowCount;
      const totalRow = sheet.addRow({
        calidad: 'TOTALES ACUMULADOS:',
        cantidad: { formula: `=SUM(E2:E${ultFila})` }
      });

      const headerRow = sheet.getRow(1);
      headerRow.height = 24;
      headerRow.eachCell(cell => {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF15803D' } };
        cell.font = { name: 'Arial', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
      });

      sheet.eachRow((row, rNum) => {
        if (rNum === 1 || rNum === ultFila + 1) return;
        row.height = 20;
        const cebra = rNum % 2 === 0;
        row.eachCell((cell, colN) => {
          cell.font = { name: 'Arial', size: 9 };
          if (cebra) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE2EFDA' } };
          if ([1, 2, 4, 6].includes(colN)) cell.alignment = { vertical: 'middle', horizontal: 'center' };
          else if (colN === 5) cell.alignment = { vertical: 'middle', horizontal: 'right' };
          else cell.alignment = { vertical: 'middle', horizontal: 'left' };
        });
      });

      totalRow.height = 22;
      totalRow.eachCell((cell, colN) => {
        cell.font = { name: 'Arial', size: 10, bold: true, color: { argb: 'FF15803D' } };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD6EEFC' } };
        cell.border = { top: { style: 'thin' }, bottom: { style: 'double' } };
        if (colN === 4 || colN === 5) cell.alignment = { vertical: 'middle', horizontal: 'right' };
      });

      sheet.autoFilter = { from: { row: 1, column: 1 }, to: { row: ultFila, column: sheet.columnCount } };

      const buffer = await workbook.xlsx.writeBuffer();
      const fechaHoy = new Date().toISOString().split('T')[0];
      const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      saveAs(blob, `REPORTE_COSECHA_CAMPO_${fechaHoy}.xlsx`);

      mostrarAlerta("Reporte de cosecha exportado a Excel con éxito", "exito");
    } catch (err) {
      console.error("Error al exportar cosecha:", err);
      mostrarAlerta("Error al generar el archivo Excel", "error");
    }
  };

  // --- 5. SEPARACIÓN DE INVERNADEROS OPERATIVOS ---
  const invernaderosOperativos = (listaInvernaderos || []).filter(i => i.activo !== false);
  const idsOperativos = invernaderosOperativos.map(i => i.id?.toString());

  // --- 6. LÓGICA DE FILTRADO Y BÚSQUEDA ---
  const registrosFiltrados = registrosCosecha.filter(item => {
    const cumpleInvernadero = !filtroInvernadero 
      ? idsOperativos.includes(item.invernadero_id?.toString()) 
      : item.invernadero_id?.toString() === filtroInvernadero.toString();
      
    const cumpleProducto = filtroProducto ? item.producto === filtroProducto : true;
    
    const q = busqueda.toLowerCase();
    const cumpleTexto = (item.operario_recolector || '').toLowerCase().includes(q) ||
                         (item.producto || '').toLowerCase().includes(q) ||
                         (item.observaciones || '').toLowerCase().includes(q);

    return cumpleInvernadero && cumpleProducto && cumpleTexto;
  });

  // A. Suma exclusiva para CANASTILLA
  const totalCanastillas = registrosFiltrados
    .filter(r => r.unidad_medida === 'CANASTILLA')
    .reduce((acc, r) => acc + parseFloat(r.cantidad || 0), 0);

  // B. Diccionario Dinámico de Totales
  const otrasUnidadesAgrupadas = {};
  registrosFiltrados.forEach(r => {
    const unidad = r.unidad_medida || 'UNIDAD';
    if (unidad !== 'CANASTILLA') {
      const cantidad = parseFloat(r.cantidad || 0);
      otrasUnidadesAgrupadas[unidad] = (otrasUnidadesAgrupadas[unidad] || 0) + cantidad;
    }
  });

  return (
    <div className="space-y-6 pb-20 text-slate-800 dark:text-slate-200 font-sans transition-colors duration-300">
      
      {/* 🚀 CABECERA PRINCIPAL MODERNA */}
      <div className="bg-white dark:bg-slate-800 p-6 rounded-3xl shadow-xl border border-slate-200 dark:border-slate-700 flex flex-col md:flex-row justify-between items-center gap-4 transition-colors duration-300">
        <div className="space-y-1">
          <div className="flex items-center gap-3">
            <span className="text-2xl p-2 bg-green-800/10 dark:bg-emerald-500/20 rounded-xl text-green-800 dark:text-emerald-400">🚜</span>
            <h2 className="text-xl font-black tracking-tight text-slate-900 dark:text-white uppercase">Planilla de Cosecha Diaria</h2>
          </div>
          <p className="text-xs text-slate-500 dark:text-slate-400 font-medium">Registro de recolección, rendimientos y operarios de campo.</p>
        </div>
        
        <div className="flex items-center gap-3 w-full md:w-auto">
          <button onClick={exportarCosechaAExcel} className="flex-1 md:flex-none px-4 py-2.5 bg-emerald-600 hover:bg-emerald-700 text-white font-black text-xs uppercase tracking-wider rounded-xl shadow-md transition-all flex items-center justify-center gap-2 cursor-pointer">
            <span>📊</span> Exportar Excel
          </button>
          <button onClick={abrirModalNuevo} className="flex-1 md:flex-none px-5 py-2.5 bg-green-800 hover:bg-green-900 text-white font-black text-xs uppercase tracking-wider rounded-xl shadow-md transition-all flex items-center justify-center gap-2 cursor-pointer">
            <span>+</span> Registrar Cosecha
          </button>
        </div>
      </div>

      {/* 🔍 BARRA DE FILTROS Y BÚSQUEDA */}
      <div className="bg-white dark:bg-slate-800 p-4 rounded-2xl shadow-md border border-slate-200 dark:border-slate-700 grid grid-cols-1 md:grid-cols-3 gap-4 transition-colors duration-300">
        <div>
          <label className="text-[10px] font-black text-slate-500 dark:text-slate-400 uppercase tracking-wider mb-1 block">🌿 Filtrar Cultivo</label>
          <select className="w-full bg-slate-50 dark:bg-slate-900 border border-slate-200 dark:border-slate-700 text-slate-800 dark:text-white p-2.5 rounded-xl font-bold text-xs uppercase outline-none focus:border-green-700" value={filtroProducto} onChange={e => setFiltroProducto(e.target.value)}>
            <option value="">TODOS LOS PRODUCTOS</option>
            {listaProductos.map(p => <option key={p.id} value={p.nombre_producto}>{p.nombre_producto}</option>)}
          </select>
        </div>

        <div>
          <label className="text-[10px] font-black text-slate-500 dark:text-slate-400 uppercase tracking-wider mb-1 block">🔍 Filtrar Ubicación</label>
          <select className="w-full bg-slate-50 dark:bg-slate-900 border border-slate-200 dark:border-slate-700 text-slate-800 dark:text-white p-2.5 rounded-xl font-bold text-xs uppercase outline-none focus:border-green-700" value={filtroInvernadero} onChange={e => setFiltroInvernadero(e.target.value)}>
            <option value="">TODA LA GRANJA (EN PRODUCCIÓN)</option>
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

        <div>
           <label className="text-[10px] font-black text-slate-500 dark:text-slate-400 uppercase tracking-wider mb-1 block">🔍 Buscar Registro</label>
           <input 
            type="text" 
            placeholder="Buscar operario, notas..." 
            className="w-full bg-slate-50 dark:bg-slate-700/50 border border-slate-200 dark:border-slate-600 text-xs text-slate-800 dark:text-white rounded-xl px-4 py-2.5 outline-none focus:border-green-700 font-bold placeholder-slate-400"
            value={busqueda}
            onChange={e => setBusqueda(e.target.value)}
          />
        </div>
      </div>

      {/* 📈 TABLEROS DE RENDIMIENTO (KPIs) */}
      <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
        <div className="bg-gradient-to-br from-green-700 to-green-900 dark:from-emerald-900 dark:to-slate-900 p-6 rounded-3xl shadow-xl text-white flex flex-col justify-between border border-green-600 dark:border-emerald-800">
          <div>
            <p className="text-[10px] font-black text-green-200 dark:text-emerald-300 uppercase tracking-widest">Rendimiento en Medida Reina</p>
            <p className="text-4xl font-black mt-2">{totalCanastillas} <span className="text-sm font-bold text-green-300 dark:text-emerald-400">Canastillas</span></p>
          </div>
          <p className="text-[10px] text-green-100 dark:text-slate-300 mt-4 font-bold border-t border-green-600/50 dark:border-slate-700 pt-3 uppercase tracking-wider">
            Total {filtroProducto || 'Global'} cosechado en Canastillas
          </p>
        </div>

        <div className="bg-gradient-to-br from-amber-600 to-amber-800 dark:from-amber-950 dark:to-slate-900 p-6 rounded-3xl shadow-xl text-white flex flex-col justify-between border border-amber-500 dark:border-amber-900">
          <div>
            <p className="text-[10px] font-black text-amber-200 dark:text-amber-400 uppercase tracking-widest mb-3">
              Desglose Lote: {filtroProducto || 'TODOS'}
            </p>
            
            <div className="space-y-2 max-h-24 overflow-y-auto pr-1 custom-scrollbar">
              {Object.keys(otrasUnidadesAgrupadas).length === 0 ? (
                <div className="text-xs font-bold text-amber-200 dark:text-slate-400 py-2">
                  No hay otras unidades registradas en este filtro.
                </div>
              ) : (
                Object.entries(otrasUnidadesAgrupadas).map(([unidadMedida, totalCantidad]) => (
                  <div key={unidadMedida} className="flex justify-between items-center border-b border-amber-500/40 dark:border-amber-900/50 pb-1.5">
                    <span className="text-xs font-black text-amber-100 dark:text-amber-300 uppercase">📦 TOTAL {unidadMedida}:</span>
                    <span className="font-black text-lg">{totalCantidad}</span>
                  </div>
                ))
              )}
            </div>
          </div>
          <p className="text-[10px] text-amber-200 dark:text-slate-400 mt-3 font-bold border-t border-amber-500/40 dark:border-amber-900/50 pt-3 uppercase tracking-wider">
            Suma automática de empaques secundarios
          </p>
        </div>
      </div>

      {/* 📊 TABLA DE HISTORIAL (ANCHO COMPLETO) */}
      <div className="bg-white dark:bg-slate-800 rounded-3xl shadow-xl overflow-hidden border border-slate-200 dark:border-slate-700 transition-colors duration-300">
        <div className="p-4 bg-slate-800 dark:bg-slate-900 text-white font-black text-xs uppercase tracking-wider flex justify-between items-center">
          <span>Historial de Recolección</span>
          <span className="bg-slate-700 dark:bg-slate-800 px-3 py-1 rounded-lg text-[10px]">{registrosFiltrados.length} Registros</span>
        </div>

        <div className="overflow-x-auto max-h-[600px]">
          <table className="w-full text-left border-collapse text-xs">
            <thead>
              <tr className="bg-slate-100 dark:bg-slate-700/80 text-slate-700 dark:text-slate-300 uppercase font-black text-[10px] tracking-wider border-b border-slate-200 dark:border-slate-700 sticky top-0 z-10">
                <th className="p-4">Fecha</th>
                <th className="p-4">Invernadero</th>
                <th className="p-4">Producto / Calidad</th>
                <th className="p-4">Recolector</th>
                <th className="p-4 text-right">Cant. Recogida</th>
                {userRole === 'admin' && <th className="p-4 text-center">Acción</th>}
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100 dark:divide-slate-700/60 font-bold text-slate-700 dark:text-slate-300">
              {registrosFiltrados.length === 0 ? (
                <tr>
                  <td colSpan={userRole === 'admin' ? 6 : 5} className="p-8 text-center text-slate-400 dark:text-slate-500 italic font-bold">
                    No hay registros de cosecha cargados para este filtro.
                  </td>
                </tr>
              ) : (
                registrosFiltrados.map((item, idx) => (
                  <tr key={item.id} className={`${idx % 2 === 0 ? 'bg-white dark:bg-slate-800' : 'bg-slate-50/50 dark:bg-slate-800/60'} hover:bg-sky-50/50 dark:hover:bg-slate-700/50 transition-colors border-l-8 border-green-700 dark:border-emerald-600`}>
                    
                    <td className="p-4 whitespace-nowrap text-slate-900 dark:text-white font-black">
                      {item.fecha_cosecha}
                    </td>
                    
                    <td className="p-4 font-black text-slate-700 dark:text-slate-300 uppercase">
                      <span className="bg-slate-200 dark:bg-slate-700 text-slate-800 dark:text-slate-200 px-2 py-1 rounded-md text-[9px] shadow-sm">
                        {item.invernaderos?.nombre || 'GENERAL'}
                      </span>
                    </td>
                    
                    <td className="p-4 uppercase">
                      <span className="font-black text-slate-900 dark:text-white block">{item.producto}</span>
                      <span className="text-[10px] font-bold text-[#117097] dark:text-sky-400 mt-0.5 inline-block">📌 {item.calidad}</span>
                    </td>
                    
                    <td className="p-4 uppercase text-emerald-800 dark:text-emerald-400 font-black">
                      <span className="flex items-center gap-1.5"><span>👤</span> {item.operario_recolector || 'N/R'}</span>
                    </td>
                    
                    <td className="p-4 text-right whitespace-nowrap">
                      <span className="font-black text-base text-green-800 dark:text-emerald-400">{item.cantidad}</span> 
                      <span className="text-[10px] text-slate-500 dark:text-slate-400 font-black ml-1 uppercase">{item.unidad_medida}</span>
                    </td>
                    
                    {userRole === 'admin' && (
                      <td className="p-4 text-center">
                        <button 
                          onClick={() => eliminarRegistroCosecha(item.id)} 
                          className="p-1.5 bg-red-100 dark:bg-red-950/60 text-red-700 dark:text-red-400 rounded-lg hover:bg-red-700 hover:text-white transition-colors border border-red-200 dark:border-red-900 cursor-pointer" 
                          title="Eliminar registro"
                        >
                          🗑️
                        </button>
                      </td>
                    )}
                  </tr>
                ))
              )}
            </tbody>
          </table>
        </div>
      </div>

      {/* 🧊 MODAL FLOTANTE PARA REGISTRO DE COSECHA */}
      {modalAbierto && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 backdrop-blur-sm p-4 overflow-y-auto">
          <div className="bg-white dark:bg-slate-800 border border-slate-200 dark:border-slate-700 w-full max-w-2xl rounded-3xl shadow-2xl overflow-hidden animate-in fade-in zoom-in-95 duration-200 my-8 text-slate-800 dark:text-slate-200">
            
            <div className="p-5 bg-slate-900 text-white border-b border-slate-800 flex justify-between items-center">
              <div className="flex items-center gap-3">
                <span className="text-xl">🚜</span>
                <div>
                  <h3 className="text-base font-black uppercase tracking-tight">Registro de Cosecha</h3>
                  <p className="text-[10px] text-slate-400 font-medium">Planilla diaria de recolección de campo.</p>
                </div>
              </div>
              <button onClick={() => setModalAbierto(false)} className="text-slate-400 hover:text-white text-lg font-black px-3 py-1 rounded-xl bg-slate-800 cursor-pointer">✕</button>
            </div>

            <form onSubmit={guardarRegistroCosecha} className="p-6 space-y-4 max-h-[80vh] overflow-y-auto">
              
              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                <div>
                  <label className="text-[10px] font-black text-slate-500 dark:text-slate-400 uppercase tracking-wider">Fecha de Corte *</label>
                  <input 
                    type="date" 
                    className="w-full bg-slate-50 dark:bg-slate-900 border border-slate-200 dark:border-slate-700 text-slate-800 dark:text-white p-2.5 rounded-xl font-bold text-sm outline-none focus:border-green-700 mt-1" 
                    value={formCosecha.fecha_cosecha} 
                    onChange={e => setFormCosecha({...formCosecha, fecha_cosecha: e.target.value})} 
                    required 
                  />
                </div>
                <div>
                  <label className="text-[10px] font-black text-slate-500 dark:text-slate-400 uppercase tracking-wider">Invernadero / Bloque *</label>
                  <select 
                    className="w-full bg-slate-50 dark:bg-slate-900 border border-slate-200 dark:border-slate-700 text-slate-800 dark:text-white p-2.5 rounded-xl font-bold text-xs uppercase outline-none focus:border-green-700 mt-1" 
                    value={formCosecha.invernadero_id} 
                    onChange={e => setFormCosecha({...formCosecha, invernadero_id: e.target.value})} 
                    required
                  >
                    <option value="">Seleccione Invernadero...</option>
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
              </div>

              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                <div>
                  <label className="text-[10px] font-black text-slate-500 dark:text-slate-400 uppercase tracking-wider">Producto *</label>
                  <select 
                    className="w-full bg-slate-50 dark:bg-slate-900 border border-slate-200 dark:border-slate-700 text-slate-800 dark:text-white p-2.5 rounded-xl font-bold text-xs uppercase outline-none focus:border-green-700 mt-1" 
                    value={formCosecha.producto} 
                    onChange={e => setFormCosecha({...formCosecha, producto: e.target.value})} 
                    required
                  >
                    {listaProductos.map(p => <option key={p.id} value={p.nombre_producto}>{p.nombre_producto}</option>)}
                  </select>
                </div>
                <div>
                  <label className="text-[10px] font-black text-slate-500 dark:text-slate-400 uppercase tracking-wider">Calidad *</label>
                  <select 
                    className="w-full bg-slate-50 dark:bg-slate-900 border border-slate-200 dark:border-slate-700 text-slate-800 dark:text-white p-2.5 rounded-xl font-bold text-xs uppercase outline-none focus:border-green-700 mt-1" 
                    value={formCosecha.calidad} 
                    onChange={e => setFormCosecha({...formCosecha, calidad: e.target.value})} 
                    required
                  >
                    {listaCalidades.map(c => <option key={c.id} value={c.nombre_calidad}>{c.nombre_calidad}</option>)}
                  </select>
                </div>
              </div>

              <div className="bg-slate-50 dark:bg-slate-900/60 p-4 rounded-2xl border border-slate-200 dark:border-slate-700 space-y-3 shadow-inner">
                <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                  <div>
                    <label className="text-[10px] font-black text-slate-500 dark:text-slate-400 uppercase tracking-wider">Cantidad Recogida *</label>
                    <input 
                      type="number" 
                      step="any" 
                      className="w-full bg-white dark:bg-slate-800 border border-slate-200 dark:border-slate-700 p-2.5 rounded-xl font-black text-lg text-green-800 dark:text-emerald-400 outline-none focus:border-green-700 mt-1 placeholder-slate-400" 
                      value={formCosecha.cantidad} 
                      onChange={e => setFormCosecha({...formCosecha, cantidad: e.target.value})} 
                      placeholder="Ej: 45" 
                      required 
                    />
                  </div>
                  <div>
                    <label className="text-[10px] font-black text-slate-500 dark:text-slate-400 uppercase tracking-wider">Unidad Medida *</label>
                    <select 
                      className="w-full bg-white dark:bg-slate-800 border border-slate-200 dark:border-slate-700 text-slate-800 dark:text-white p-2.5 rounded-xl font-bold text-xs uppercase outline-none focus:border-green-700 mt-1" 
                      value={formCosecha.unidad_medida} 
                      onChange={e => setFormCosecha({...formCosecha, unidad_medida: e.target.value})} 
                      required
                    >
                      {listaUnidades.map(u => <option key={u.id} value={u.nombre_unidad}>{u.nombre_unidad}</option>)}
                    </select>
                  </div>
                </div>
              </div>

              <div>
                <label className="text-[10px] font-black text-slate-500 dark:text-slate-400 uppercase tracking-wider">Operario Encargado / Recolector</label>
                <select 
                  className="w-full bg-slate-50 dark:bg-slate-900 border border-slate-200 dark:border-slate-700 text-slate-800 dark:text-white p-2.5 rounded-xl font-bold text-xs uppercase outline-none focus:border-green-700 mt-1" 
                  value={formCosecha.operario_recolector} 
                  onChange={e => setFormCosecha({...formCosecha, operario_recolector: e.target.value})}
                >
                  <option value="">Seleccione operario recolector...</option>
                  {listaTrabajadores.map(trab => (
                    <option key={trab.id} value={trab.nombre_completo}>{trab.nombre_completo}</option>
                  ))}
                </select>
              </div>

              <div>
                <label className="text-[10px] font-black text-slate-500 dark:text-slate-400 uppercase tracking-wider">Observaciones del Lote</label>
                <textarea 
                  className="w-full bg-slate-50 dark:bg-slate-900 border border-slate-200 dark:border-slate-700 text-slate-800 dark:text-white p-2.5 rounded-xl font-bold h-20 text-xs uppercase outline-none focus:border-green-700 mt-1 resize-none placeholder-slate-400" 
                  value={formCosecha.observaciones} 
                  onChange={e => setFormCosecha({...formCosecha, observaciones: e.target.value})} 
                  placeholder="Ej: Sin rastro de plaga" 
                />
              </div>

              <div className="flex justify-end gap-3 pt-4 border-t border-slate-200 dark:border-slate-700">
                <button type="button" onClick={() => setModalAbierto(false)} className="px-5 py-2.5 bg-slate-200 dark:bg-slate-700 hover:bg-slate-300 dark:hover:bg-slate-600 text-slate-700 dark:text-slate-200 font-black text-xs uppercase tracking-wider rounded-xl transition-colors cursor-pointer">
                  Cancelar
                </button>
                <button type="submit" className="px-6 py-2.5 bg-green-800 hover:bg-green-900 text-white font-black text-xs uppercase tracking-wider rounded-xl shadow-lg transition-colors cursor-pointer">
                  💾 Registrar Cosecha
                </button>
              </div>

            </form>
          </div>
        </div>
      )}

    </div>
  );
}