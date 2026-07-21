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
    <div className="space-y-6 pb-20 text-slate-800">
      
      {/* SECCIÓN SUPERIOR DE FILTROS ORGANIZADOS POR GRUPOS */}
      <div className="bg-slate-100 p-4 rounded-2xl border border-slate-200 grid grid-cols-1 md:grid-cols-2 gap-4">
        <div>
          <label className="text-[10px] font-black text-slate-500 uppercase tracking-wider mb-1 block">🌿 1. Filtrar Cultivo / Producto</label>
          <select className="w-full border-2 p-2.5 rounded-xl font-black bg-white border-slate-300 text-xs uppercase outline-none focus:border-green-700" value={filtroProducto} onChange={e => setFiltroProducto(e.target.value)}>
            <option value="">VER TODOS LOS PRODUCTOS</option>
            {listaProductos.map(p => <option key={p.id} value={p.nombre_producto}>{p.nombre_producto}</option>)}
          </select>
        </div>

        <div>
          <label className="text-[10px] font-black text-slate-500 uppercase tracking-wider mb-1 block">🔍 2. Filtrar Ubicación / Invernadero</label>
          <select className="w-full border-2 p-2.5 rounded-xl font-black bg-white border-slate-300 text-xs uppercase outline-none focus:border-green-700" value={filtroInvernadero} onChange={e => setFiltroInvernadero(e.target.value)}>
            <option value="">VER TODA LA GRANJA (EN PRODUCCIÓN)</option>
            
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

      {/* TABLEROS DE RENDIMIENTO */}
      <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
        <div className="bg-gradient-to-br from-green-700 to-green-900 p-5 rounded-2xl shadow-xl text-white flex flex-col justify-between">
          <div>
            <p className="text-[10px] font-black text-green-200 uppercase tracking-widest italic">Rendimiento en Medida Reina</p>
            <p className="text-4xl font-black mt-2">{totalCanastillas} <span className="text-xs font-bold text-green-300">Canastillas</span></p>
          </div>
          <p className="text-[10px] text-green-100 mt-4 font-bold border-t border-green-600/50 pt-2 uppercase tracking-tighter">
            Total {filtroProducto || 'Global'} cosechado en Canastillas
          </p>
        </div>

        <div className="bg-gradient-to-br from-amber-600 to-amber-800 p-5 rounded-2xl shadow-xl text-white flex flex-col justify-between">
          <div>
            <p className="text-[10px] font-black text-amber-200 uppercase tracking-widest italic mb-3">
              Desglose Lote: {filtroProducto || 'TODOS'}
            </p>
            
            <div className="space-y-2 max-h-32 overflow-y-auto pr-1">
              {Object.keys(otrasUnidadesAgrupadas).length === 0 ? (
                <div className="text-xs font-bold text-amber-200 italic py-2">
                  No hay otras unidades registradas en este filtro.
                </div>
              ) : (
                Object.entries(otrasUnidadesAgrupadas).map(([unidadMedida, totalCantidad]) => (
                  <div key={unidadMedida} className="flex justify-between items-center border-b border-amber-500/40 pb-1.5">
                    <span className="text-xs font-black text-amber-100 uppercase">📦 TOTAL {unidadMedida}:</span>
                    <span className="font-black text-lg">{totalCantidad}</span>
                  </div>
                ))
              )}
            </div>
          </div>
          <p className="text-[10px] text-amber-200 mt-2 font-bold border-t border-amber-500/40 pt-2 uppercase tracking-tighter">
            Suma automática de empaques secundarios
          </p>
        </div>
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
        {/* FORMULARIO DE RECOLECCIÓN DIARIA */}
        <div className="bg-white p-6 rounded-3xl shadow-xl border-t-8 border-green-800 h-fit">
          <h3 className="font-black text-slate-800 uppercase text-xs italic mb-5">🚜 Planilla de Cosecha Diaria</h3>
          <form onSubmit={guardarRegistroCosecha} className="space-y-4">
            <div>
              <label className="text-[10px] font-bold text-gray-400 uppercase px-1 italic">Fecha de Corte *</label>
              <input type="date" className="w-full border-2 p-2.5 rounded-xl font-bold text-sm outline-none focus:border-green-700" value={formCosecha.fecha_cosecha} onChange={e => setFormCosecha({...formCosecha, fecha_cosecha: e.target.value})} required />
            </div>

            <div>
              <label className="text-[10px] font-bold text-gray-400 uppercase px-1 italic">Invernadero / Bloque *</label>
              <select className="w-full border-2 p-2.5 rounded-xl font-bold bg-white text-xs outline-none focus:border-green-700" value={formCosecha.invernadero_id} onChange={e => setFormCosecha({...formCosecha, invernadero_id: e.target.value})} required>
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

            <div className="grid grid-cols-2 gap-2">
              <div>
                <label className="text-[10px] font-bold text-gray-400 uppercase px-1 italic">Producto *</label>
                <select className="w-full border-2 p-2.5 rounded-xl font-bold bg-white text-xs uppercase outline-none focus:border-green-700" value={formCosecha.producto} onChange={e => setFormCosecha({...formCosecha, producto: e.target.value})} required>
                  {listaProductos.map(p => <option key={p.id} value={p.nombre_producto}>{p.nombre_producto}</option>)}
                </select>
              </div>
              <div>
                <label className="text-[10px] font-bold text-gray-400 uppercase px-1 italic">Calidad *</label>
                <select className="w-full border-2 p-2.5 rounded-xl font-bold bg-white text-xs uppercase outline-none focus:border-green-700" value={formCosecha.calidad} onChange={e => setFormCosecha({...formCosecha, calidad: e.target.value})} required>
                  {listaCalidades.map(c => <option key={c.id} value={c.nombre_calidad}>{c.nombre_calidad}</option>)}
                </select>
              </div>
            </div>

            <div className="grid grid-cols-2 gap-2">
              <div>
                <label className="text-[10px] font-bold text-gray-400 uppercase px-1 italic">Cantidad Recogida *</label>
                <input type="number" step="any" className="w-full border-2 p-2.5 rounded-xl font-black text-lg text-green-800 outline-none focus:border-green-700" value={formCosecha.cantidad} onChange={e => setFormCosecha({...formCosecha, cantidad: e.target.value})} placeholder="Ej: 45" required />
              </div>
              <div>
                <label className="text-[10px] font-bold text-gray-400 uppercase px-1 italic">Unidad Medida *</label>
                <select className="w-full border-2 p-2.5 rounded-xl font-bold bg-white text-xs uppercase outline-none focus:border-green-700" value={formCosecha.unidad_medida} onChange={e => setFormCosecha({...formCosecha, unidad_medida: e.target.value})} required>
                  {listaUnidades.map(u => <option key={u.id} value={u.nombre_unidad}>{u.nombre_unidad}</option>)}
                </select>
              </div>
            </div>

            <div>
              <label className="text-[10px] font-bold text-gray-400 uppercase px-1 italic">Operario Encargado / Recolector (Nómina)</label>
              <select 
                className="w-full border-2 p-2.5 rounded-xl font-bold uppercase bg-white text-xs outline-none focus:border-green-700" 
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
              <label className="text-[10px] font-bold text-gray-400 uppercase px-1 italic">Observaciones del Lote</label>
              <textarea className="w-full border-2 p-2.5 rounded-xl font-bold h-16 bg-gray-50 uppercase focus:bg-white text-xs outline-none focus:border-green-700" value={formCosecha.observaciones} onChange={e => setFormCosecha({...formCosecha, observaciones: e.target.value})} placeholder="Ej: Sin rastro de plaga" />
            </div>

            <button type="submit" className="w-full py-3.5 bg-green-800 text-white font-black rounded-xl uppercase text-xs tracking-widest hover:bg-green-900 shadow-md transition-colors cursor-pointer">💾 Registrar Cosecha</button>
          </form>
        </div>

        {/* TABLA HISTÓRICA */}
        <div className="lg:col-span-2 bg-white rounded-3xl shadow-xl overflow-hidden border border-gray-200 flex flex-col">
          
          <div className="p-4 bg-slate-800 text-white font-black text-xs uppercase tracking-wider flex flex-col sm:flex-row justify-between items-center gap-3">
            <span>Historial de Recolección ({registrosFiltrados.length})</span>

            <div className="flex items-center gap-2 flex-wrap w-full sm:w-auto justify-end">
              <button
                onClick={exportarCosechaAExcel}
                className="px-3 py-1.5 bg-emerald-600 hover:bg-emerald-700 text-white text-[10px] font-black uppercase rounded-xl shadow transition-colors flex items-center gap-1 cursor-pointer"
              >
                📊 EXPORTAR EXCEL
              </button>

              <input 
                type="text" 
                placeholder="🔍 Buscar por operario, notas..." 
                className="px-3 py-1.5 text-xs rounded-xl text-slate-800 outline-none font-bold placeholder-gray-400 min-w-[180px]"
                value={busqueda}
                onChange={e => setBusqueda(e.target.value)}
              />
            </div>
          </div>

          <div className="overflow-x-auto min-h-[400px]">
            <table className="w-full text-left border-collapse text-[11px]">
              <thead>
                <tr className="bg-gray-200 text-slate-800 uppercase font-black sticky top-0">
                  <th className="p-3">Fecha</th>
                  <th className="p-3">Invernadero</th>
                  <th className="p-3">Producto / Calidad</th>
                  <th className="p-3">Recolector</th>
                  <th className="p-3 text-right">Cant. Recogida</th>
                  {userRole === 'admin' && <th className="p-3 text-center">Acción</th>}
                </tr>
              </thead>
              <tbody className="divide-y divide-gray-200 font-bold text-slate-700">
                {registrosFiltrados.length === 0 ? (
                  <tr><td colSpan={userRole === 'admin' ? 6 : 5} className="p-8 text-center text-gray-400 italic font-bold">No hay registros de cosecha cargados para este filtro</td></tr>
                ) : (
                  registrosFiltrados.map((item, idx) => (
                    <tr key={item.id} className={`${idx % 2 === 0 ? 'bg-white' : 'bg-slate-50'} hover:bg-sky-50 transition-colors`}>
                      <td className="p-3 whitespace-nowrap text-slate-600 border-l-4 border-green-700">{item.fecha_cosecha}</td>
                      <td className="p-3 font-black text-slate-900 uppercase">{item.invernaderos?.nombre || 'GENERAL'}</td>
                      <td className="p-3 uppercase">
                        <span className="font-black text-slate-900 block">{item.producto}</span>
                        <span className="text-[9px] font-bold text-gray-400">{item.calidad}</span>
                      </td>
                      <td className="p-3 uppercase text-emerald-800 font-black">
                        👤 {item.operario_recolector || 'N/R'}
                      </td>
                      <td className="p-3 text-right font-black text-sm text-green-800">
                        {item.cantidad} <span className="text-[10px] text-gray-500 font-normal">{item.unidad_medida}</span>
                      </td>
                      {userRole === 'admin' && (
                        <td className="p-3 text-center">
                          <button 
                            onClick={() => eliminarRegistroCosecha(item.id)} 
                            className="p-1 bg-red-100 text-red-600 rounded-lg hover:bg-red-600 hover:text-white transition-all text-xs" 
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
      </div>
    </div>
  );
}