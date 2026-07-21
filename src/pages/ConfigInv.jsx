import React, { useState, useEffect } from 'react';
import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';

export default function ConfigInv({ invForm, setInvForm, mostrarAlerta, cargarTodo, supabase, lista }) {
  
  const [listaProductos, setListaProductos] = useState([]);
  const [listaCalidades, setListaCalidades] = useState([]); 
  const [tabTabla, setTabTabla] = useState('activos'); // 'activos' o 'archivados'
  const [busqueda, setBusqueda] = useState('');

  useEffect(() => {
    obtenerParametrosConfigurados();
  }, []);

  const obtenerParametrosConfigurados = async () => {
    try {
      const [resProd, resCal] = await Promise.all([
        supabase.from('config_productos').select('*').order('nombre_producto', { ascending: true }),
        supabase.from('config_calidades').select('*').order('nombre_calidad', { ascending: true }) 
      ]);

      if (resProd.error) throw resProd.error;
      if (resCal.error) throw resCal.error;
      
      const productosObtenidos = resProd.data || [];
      const calidadesObtenidas = resCal.data || [];

      setListaProductos(productosObtenidos);
      setListaCalidades(calidadesObtenidas);

      setInvForm(prev => ({
        ...prev,
        cultivo: prev.cultivo || productosObtenidos[0]?.nombre_producto || '',
        variedad: prev.variedad || calidadesObtenidas[0]?.nombre_calidad || '' 
      }));

    } catch (err) {
      console.error("Error obteniendo parámetros para invernaderos:", err);
      mostrarAlerta("No se pudieron cargar las configuraciones", "error");
    }
  };

  const prepararEdicionInvernadero = (item) => {
    let estadoDB = String(item.state || item.estado || 'ACTIVO')
      .toUpperCase()
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .trim();

    let estadoLimpio = 'ACTIVO';
    if (estadoDB.includes("COSECHA")) {
      estadoLimpio = "EN_COSECHA";
    } else if (estadoDB.includes("PREPARAC")) {
      estadoLimpio = "EN_PREPARACION";
    } else if (estadoDB.includes("DESCANSO") || (estadoDB === "INACTIVO" && item.activo !== false)) {
      estadoLimpio = "EN_DESCANSO";
    } else if (item.activo === false) {
      estadoLimpio = "INACTIVO";
    }

    setInvForm({
      id_editando: item.id, 
      nombre: item.nombre,
      cultivo: item.cultivo || item.cultivo_principal || '',
      variedad: item.variedad || '',
      largo: item.largo || '',
      ancho: item.ancho || '',
      siembra: item.fecha_siembra || '',
      cosecha: item.fecha_cosecha_est || '',
      estado: estadoLimpio, 
      descripcion: item.descripcion || ''
    });
    window.scrollTo({ top: 0, behavior: 'smooth' }); 
  };

  const limpiarFormulario = () => {
    setInvForm({
      id_editando: null,
      nombre: '', 
      cultivo: listaProductos[0]?.nombre_producto || '', 
      variedad: listaCalidades[0]?.nombre_calidad || '', 
      largo: '', 
      ancho: '',
      siembra: '', 
      cosecha: '', 
      estado: 'ACTIVO', 
      descripcion: ''
    });
  };

  const handleSave = async () => {
    if (!invForm.nombre) {
      mostrarAlerta("El nombre es obligatorio", "error");
      return;
    }

    const payload = {
      nombre: invForm.nombre.toUpperCase().trim(),
      cultivo_principal: invForm.cultivo || null,
      variedad: invForm.variedad || null, 
      largo: parseFloat(invForm.largo) || 0,
      ancho: parseFloat(invForm.ancho) || 0,
      fecha_siembra: invForm.siembra || null,
      fecha_cosecha_est: invForm.cosecha || null, 
      estado: invForm.estado, 
      descripcion: invForm.descripcion ? invForm.descripcion.toUpperCase().trim() : null,
      cultivo: invForm.cultivo || null,
      activo: true
    };

    try {
      let error;
      if (invForm.id_editando) {
        const res = await supabase.from('invernaderos').update(payload).eq('id', invForm.id_editando);
        error = res.error;
      } else {
        const res = await supabase.from('invernaderos').insert([payload]);
        error = res.error;
      }

      if (error) throw error;

      mostrarAlerta(invForm.id_editando ? "Invernadero actualizado exitosamente" : "Invernadero creado exitosamente", "exito");
      limpiarFormulario();
      cargarTodo();
    } catch (err) {
      console.error("Error al procesar invernadero:", err);
      mostrarAlerta("Error de envío: " + err.message, "error");
    }
  };

  // ARCHIVAR / INACTIVAR (LÓGICO)
  const handleInactivarLogico = async (id, nombre) => {
    const confirmar = window.confirm(`¿Estás seguro de ARCHIVAR el "${nombre}"?\n\nToda su información contable, ventas y cosechas pasadas PERMANECERÁ GUARDADA intacta en el sistema, pero el bloque se moverá a la pestaña de "Archivados".`);
    
    if (!confirmar) return;

    try {
      const { error } = await supabase
        .from('invernaderos')
        .update({ activo: false, estado: 'INACTIVO' })
        .eq('id', id);

      if (error) throw error;

      mostrarAlerta("Invernadero archivado exitosamente. Historial conservado intacto.", "exito");
      await cargarTodo(); 
    } catch (err) {
      console.error("Error al inactivar el invernadero:", err);
      mostrarAlerta("No se pudo inactivar: " + (err.message || err), "error");
    }
  };

  // REACTIVAR INVERNADERO ARCHIVADO
  const handleReactivar = async (id, nombre) => {
    if (!window.confirm(`¿Deseas REACTIVAR el "${nombre}" para volver a operarlo?`)) return;

    try {
      const { error } = await supabase
        .from('invernaderos')
        .update({ activo: true, estado: 'ACTIVO' })
        .eq('id', id);

      if (error) throw error;

      mostrarAlerta("Invernadero reactivado y disponible para producción", "exito");
      await cargarTodo(); 
    } catch (err) {
      mostrarAlerta("Error al reactivar el invernadero", "error");
    }
  };

  // EXPORTAR CATALOGO DE INVERNADEROS A EXCEL
  const exportarInvernaderosAExcel = async () => {
    if (!lista || lista.length === 0) {
      mostrarAlerta("No hay invernaderos para exportar", "error");
      return;
    }

    try {
      const workbook = new ExcelJS.Workbook();
      const sheet = workbook.addWorksheet('Estructura Invernaderos');

      sheet.columns = [
        { header: 'INVERNADERO / BLOQUE', key: 'nombre', width: 28 },
        { header: 'CULTIVO PRINCIPAL', key: 'cultivo', width: 22 },
        { header: 'VARIEDAD', key: 'variedad', width: 18 },
        { header: 'LARGO (M)', key: 'largo', width: 14 },
        { header: 'ANCHO (M)', key: 'ancho', width: 14 },
        { header: 'ÁREA APROX (M²)', key: 'area', width: 18 },
        { header: 'FECHA SIEMBRA', key: 'siembra', width: 16 },
        { header: 'EST. COSECHA', key: 'cosecha', width: 16 },
        { header: 'ESTADO OPERATIVO', key: 'estado', width: 20 },
        { header: 'SITUACIÓN', key: 'activo', width: 16 }
      ];

      lista.forEach(inv => {
        const l = parseFloat(inv.largo || 0);
        const a = parseFloat(inv.ancho || 0);
        sheet.addRow({
          nombre: (inv.nombre || '').toUpperCase(),
          cultivo: (inv.cultivo || inv.cultivo_principal || 'S/C').toUpperCase(),
          variedad: (inv.variedad || 'S/V').toUpperCase(),
          largo: l,
          ancho: a,
          area: l * a,
          siembra: inv.fecha_siembra || 'N/R',
          cosecha: inv.fecha_cosecha_est || 'N/R',
          estado: (inv.estado || 'ACTIVO').replace('_', ' ').toUpperCase(),
          activo: inv.activo !== false ? 'OPERATIVO' : 'ARCHIVADO'
        });
      });

      const headerRow = sheet.getRow(1);
      headerRow.height = 24;
      headerRow.eachCell(cell => {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF15803D' } };
        cell.font = { name: 'Arial', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
      });

      sheet.eachRow((row, rNum) => {
        if (rNum === 1) return;
        row.height = 20;
        const cebra = rNum % 2 === 0;
        row.eachCell((cell, colN) => {
          cell.font = { name: 'Arial', size: 9 };
          if (cebra) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE2EFDA' } };
          if ([4, 5, 6].includes(colN)) cell.alignment = { vertical: 'middle', horizontal: 'right' };
          else if ([7, 8, 9, 10].includes(colN)) cell.alignment = { vertical: 'middle', horizontal: 'center' };
          else cell.alignment = { vertical: 'middle', horizontal: 'left' };
        });
      });

      sheet.autoFilter = { from: { row: 1, column: 1 }, to: { row: sheet.rowCount, column: sheet.columnCount } };

      const buffer = await workbook.xlsx.writeBuffer();
      const fechaHoy = new Date().toISOString().split('T')[0];
      const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      saveAs(blob, `ESTRUCTURA_INVERNADEROS_${fechaHoy}.xlsx`);

      mostrarAlerta("Estructura de invernaderos exportada a Excel con éxito", "exito");
    } catch (err) {
      console.error("Error al exportar invernaderos:", err);
      mostrarAlerta("Error al generar el archivo Excel", "error");
    }
  };

  // Filtrado de la lista por pestaña y búsqueda
  const listaActivos = (lista || []).filter(i => i.activo !== false);
  const listaArchivados = (lista || []).filter(i => i.activo === false);

  const listaFiltradaTabla = (tabTabla === 'activos' ? listaActivos : listaArchivados).filter(inv => {
    const q = busqueda.toLowerCase();
    return (
      (inv.nombre || '').toLowerCase().includes(q) ||
      (inv.cultivo || inv.cultivo_principal || '').toLowerCase().includes(q) ||
      (inv.variedad || '').toLowerCase().includes(q)
    );
  });

  return (
    <div className="space-y-6 pb-20 text-slate-800">
      <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
        
        {/* COLUMNA IZQUIERDA: FORMULARIO */}
        <div className="bg-white p-6 rounded-3xl shadow-xl border-t-8 border-green-700 h-fit">
          <div className="flex justify-between items-center mb-5 border-b pb-2">
            <h3 className="font-black text-slate-800 uppercase text-xs italic">
              {invForm.id_editando ? '📝 Editar Invernadero' : '🏠 Registrar Nuevo Invernadero'}
            </h3>
            {invForm.id_editando && (
              <button 
                onClick={limpiarFormulario}
                className="text-gray-400 hover:text-red-500 font-black text-sm"
                title="Cancelar edición"
              >
                ✕
              </button>
            )}
          </div>
          
          <div className="space-y-4">
            <div>
              <label className="text-[10px] font-bold text-gray-400 uppercase px-1 italic">Nombre / Identificador *</label>
              <input 
                type="text" 
                className="w-full border-2 p-2.5 rounded-xl font-bold text-sm uppercase outline-none focus:border-green-700" 
                value={invForm.nombre || ''} 
                onChange={e => setInvForm({...invForm, nombre: e.target.value})} 
                placeholder="Ej: INVERNADERO 1" 
                required 
              />
            </div>

            <div className="grid grid-cols-2 gap-2">
              <div>
                <label className="text-[10px] font-bold text-gray-400 uppercase px-1 italic">Tipo de Cultivo</label>
                <select 
                  className="w-full border-2 p-2.5 rounded-xl font-bold bg-white text-xs uppercase outline-none focus:border-green-700" 
                  value={invForm.cultivo || ''} 
                  onChange={e => setInvForm({...invForm, cultivo: e.target.value})}
                >
                  {listaProductos.length === 0 ? (
                    <option value="">No hay cultivos creados...</option>
                  ) : (
                    listaProductos.map(p => (
                      <option key={p.id} value={p.nombre_producto}>{p.nombre_producto}</option>
                    ))
                  )}
                </select>
              </div>

              <div>
                <label className="text-[10px] font-bold text-gray-400 uppercase px-1 italic">Clasificación / Variedad</label>
                <select 
                  className="w-full border-2 p-2.5 rounded-xl font-bold bg-white text-xs uppercase outline-none focus:border-green-700" 
                  value={invForm.variedad || ''} 
                  onChange={e => setInvForm({...invForm, variedad: e.target.value})}
                >
                  {listaCalidades.length === 0 ? (
                    <option value="">No hay calidades creadas...</option>
                  ) : (
                    listaCalidades.map(c => (
                      <option key={c.id} value={c.nombre_calidad}>{c.nombre_calidad}</option>
                    ))
                  )}
                </select>
              </div>
            </div>

            <div className="grid grid-cols-2 gap-2">
              <div>
                <label className="text-[10px] font-bold text-gray-400 uppercase px-1 italic">Largo (Metros)</label>
                <input 
                  type="number" 
                  className="w-full border-2 p-2.5 rounded-xl font-bold text-sm outline-none focus:border-green-700" 
                  value={invForm.largo || ''} 
                  onChange={e => setInvForm({...invForm, largo: e.target.value})} 
                  placeholder="0.00" 
                />
              </div>
              <div>
                <label className="text-[10px] font-bold text-gray-400 uppercase px-1 italic">Ancho (Metros)</label>
                <input 
                  type="number" 
                  className="w-full border-2 p-2.5 rounded-xl font-bold text-sm outline-none focus:border-green-700" 
                  value={invForm.ancho || ''} 
                  onChange={e => setInvForm({...invForm, ancho: e.target.value})} 
                  placeholder="0.00" 
                />
              </div>
            </div>

            <div className="grid grid-cols-2 gap-2">
              <div>
                <label className="text-[10px] font-bold text-gray-400 uppercase px-1 italic">Fecha Siembra</label>
                <input 
                  type="date" 
                  className="w-full border-2 p-2.5 rounded-xl font-bold text-xs outline-none focus:border-green-700" 
                  value={invForm.siembra || ''} 
                  onChange={e => setInvForm({...invForm, siembra: e.target.value})} 
                />
              </div>
              <div>
                <label className="text-[10px] font-bold text-gray-400 uppercase px-1 italic">Est. Cosecha</label>
                <input 
                  type="date" 
                  className="w-full border-2 p-2.5 rounded-xl font-bold text-xs outline-none focus:border-green-700" 
                  value={invForm.cosecha || ''} 
                  onChange={e => setInvForm({...invForm, cosecha: e.target.value})} 
                />
              </div>
            </div>

            <div>
              <label className="text-[10px] font-bold text-gray-400 uppercase px-1 italic">Estado Actual del Bloque</label>
              <select 
                className="w-full border-2 p-2.5 rounded-xl font-black bg-white text-xs outline-none focus:border-green-700 uppercase text-slate-800" 
                value={invForm.estado || 'ACTIVO'} 
                onChange={e => setInvForm({...invForm, estado: e.target.value})}
              >
                <option value="ACTIVO">ACTIVO</option>
                <option value="EN_COSECHA">EN COSECHA</option>
                <option value="EN_PREPARACION">EN PREPARACIÓN</option>
                <option value="EN_DESCANSO">EN DESCANSO</option>
              </select>
            </div>

            <div>
              <label className="text-[10px] font-bold text-gray-400 uppercase px-1 italic">Notas / Descripción</label>
              <textarea 
                className="w-full border-2 p-2.5 rounded-xl font-bold h-16 text-xs uppercase outline-none focus:border-green-700" 
                value={invForm.descripcion || ''} 
                onChange={e => setInvForm({...invForm, descripcion: e.target.value})} 
                placeholder="Ej: Suelo tratado con micorrizas" 
              />
            </div>

            <button 
              onClick={handleSave} 
              className={`w-full py-3.5 text-white font-black rounded-xl uppercase text-xs tracking-widest transition-colors shadow-md ${
                invForm.id_editando ? 'bg-amber-600 hover:bg-amber-700' : 'bg-green-800 hover:bg-green-900'
              }`}
            >
              {invForm.id_editando ? '💾 Actualizar Invernadero' : '💾 Crear Invernadero'}
            </button>
          </div>
        </div>

        {/* COLUMNA DERECHA: TABLA CON PESTAÑAS Y BÚSQUEDA */}
        <div className="lg:col-span-2 bg-white rounded-3xl shadow-xl overflow-hidden border border-gray-200 flex flex-col">
          
          <div className="p-4 bg-slate-800 text-white font-black text-xs uppercase tracking-wider flex flex-col sm:flex-row justify-between items-center gap-3">
            
            {/* PESTAÑAS: OPERATIVOS vs ARCHIVADOS */}
            <div className="flex items-center gap-2">
              <button 
                onClick={() => setTabTabla('activos')} 
                className={`px-3 py-1.5 rounded-lg text-[10px] uppercase font-black transition-all ${
                  tabTabla === 'activos' ? 'bg-green-700 text-white shadow' : 'bg-slate-700 text-slate-300'
                }`}
              >
                🌱 Operativos ({listaActivos.length})
              </button>
              <button 
                onClick={() => setTabTabla('archivados')} 
                className={`px-3 py-1.5 rounded-lg text-[10px] uppercase font-black transition-all ${
                  tabTabla === 'archivados' ? 'bg-red-600 text-white shadow' : 'bg-slate-700 text-slate-300'
                }`}
              >
                📁 Archivados / Inactivos ({listaArchivados.length})
              </button>
            </div>

            <div className="flex items-center gap-2 flex-wrap w-full sm:w-auto justify-end">
              <button
                onClick={exportarInvernaderosAExcel}
                className="px-3 py-1.5 bg-emerald-600 hover:bg-emerald-700 text-white text-[10px] font-black uppercase rounded-xl shadow transition-colors flex items-center gap-1"
              >
                📊 EXPORTAR EXCEL
              </button>

              <input 
                type="text" 
                placeholder="🔍 Buscar bloque, cultivo..." 
                className="px-3 py-1.5 text-xs rounded-xl text-slate-800 outline-none font-bold placeholder-gray-400 min-w-[180px]"
                value={busqueda}
                onChange={e => setBusqueda(e.target.value)}
              />
            </div>
          </div>

          <div className="overflow-x-auto min-h-[400px]">
            <table className="w-full text-left text-[11px] border-collapse">
              <thead>
                <tr className="bg-gray-200 text-slate-800 uppercase font-black sticky top-0">
                  <th className="p-3.5 border-b border-gray-300">Invernadero / Bloque</th>
                  <th className="p-3.5 border-b border-gray-300">Cultivo / Clasificación</th>
                  <th className="p-3.5 border-b border-gray-300 text-center">Estado</th>
                  <th className="p-3.5 border-b border-gray-300">Siembra / Est. Cosecha</th>
                  <th className="p-3.5 border-b border-gray-300 text-center">Acciones</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-gray-200 font-bold text-slate-700">
                {listaFiltradaTabla.length === 0 ? (
                  <tr>
                    <td colSpan="5" className="p-8 text-center text-gray-400 italic font-bold">
                      {tabTabla === 'activos' ? 'No hay invernaderos activos en producción.' : 'No hay invernaderos archivados en el histórico.'}
                    </td>
                  </tr>
                ) : (
                  listaFiltradaTabla.map((item, index) => {
                    const esInactivo = item.activo === false;
                    const estVisual = String(item.estado || '').toUpperCase();

                    return (
                      <tr key={item.id} className={`${index % 2 === 0 ? 'bg-white' : 'bg-slate-50'} hover:bg-sky-50 transition-colors`}>
                        
                        <td className={`p-3.5 font-black text-slate-900 border-l-4 ${esInactivo ? 'border-red-600 bg-red-50/40' : 'border-green-700'}`}>
                          <p className="uppercase text-xs">{item.nombre}</p>
                          {esInactivo && <span className="text-[9px] text-red-500 font-bold italic block mt-0.5">📁 Registros Conservados</span>}
                        </td>
                        
                        <td className="p-3.5 font-bold text-slate-700 uppercase">
                          <p>{item.cultivo || item.cultivo_principal || 'S/C'} - {item.variedad || 'S/C'}</p>
                          <p className="text-[9px] text-slate-400 font-bold lowercase italic mt-0.5">📐 Área: {item.largo || 0}m × {item.ancho || 0}m ({(parseFloat(item.largo || 0) * parseFloat(item.ancho || 0)).toLocaleString()} m²)</p>
                        </td>
                        
                        <td className="p-3.5 text-center">
                          {esInactivo ? (
                            <span className="px-2.5 py-1 rounded-md text-[9px] font-black uppercase shadow-sm bg-red-100 text-red-700 border border-red-300">
                              ARCHIVADO
                            </span>
                          ) : (
                            <span className={`px-2.5 py-1 rounded-md text-[9px] font-black uppercase shadow-sm ${
                              estVisual === 'ACTIVO' ? 'bg-green-100 text-green-700 border border-green-300' : 
                              estVisual === 'EN_COSECHA' ? 'bg-blue-100 text-blue-700 border border-blue-300' :
                              estVisual === 'EN_PREPARACION' ? 'bg-amber-100 text-amber-700 border border-amber-300' : 
                              estVisual === 'EN_DESCANSO' || estVisual === 'INACTIVO' ? 'bg-purple-100 text-purple-700 border border-purple-300' :
                              'bg-gray-200 text-gray-700 border border-gray-300'
                            }`}>
                              {estVisual === 'EN_DESCANSO' || (estVisual === 'INACTIVO' && !esInactivo) ? 'EN DESCANSO' : estVisual.replace('_', ' ')}
                            </span>
                          )}
                        </td>

                        <td className="p-3.5 font-bold text-slate-600">
                          <p>🌱 <span className="text-slate-400 text-[10px]">S:</span> {item.fecha_siembra || 'N/R'}</p>
                          <p>🚜 <span className="text-slate-400 text-[10px]">C:</span> {item.fecha_cosecha_est || 'N/R'}</p>
                        </td>

                        <td className="p-3.5 text-center">
                          <div className="flex gap-1.5 justify-center">
                            <button
                              type="button"
                              onClick={() => prepararEdicionInvernadero(item)}
                              className="p-1.5 bg-amber-100 text-amber-700 hover:bg-amber-600 hover:text-white rounded-lg border border-amber-200 transition-all text-xs"
                              title="Editar Invernadero"
                            >
                              ✏️
                            </button>
                            
                            {!esInactivo ? (
                              <button
                                type="button"
                                onClick={() => handleInactivarLogico(item.id, item.nombre)}
                                className="p-1.5 bg-red-100 text-red-700 hover:bg-red-600 hover:text-white rounded-lg border border-red-200 transition-all text-xs"
                                title="Archivar Invernadero"
                              >
                                🚫
                              </button>
                            ) : (
                              <button
                                type="button"
                                onClick={() => handleReactivar(item.id, item.nombre)}
                                className="p-1.5 bg-emerald-100 text-emerald-700 hover:bg-emerald-600 hover:text-white rounded-lg border border-emerald-300 transition-all text-xs"
                                title="Reactivar Invernadero"
                              >
                                ♻️
                              </button>
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

      </div>
    </div>
  );
}