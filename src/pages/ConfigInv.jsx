import React, { useState, useEffect } from 'react';
import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';

export default function ConfigInv({ invForm, setInvForm, mostrarAlerta, cargarTodo, supabase, lista }) {
  
  const [listaProductos, setListaProductos] = useState([]);
  const [listaCalidades, setListaCalidades] = useState([]); 
  const [tabTabla, setTabTabla] = useState('activos'); // 'activos' o 'archivados'
  const [busqueda, setBusqueda] = useState('');
  const [modalAbierto, setModalAbierto] = useState(false);

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
    setModalAbierto(true);
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

  const abrirModalNuevo = () => {
    limpiarFormulario();
    setModalAbierto(true);
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
      setModalAbierto(false);
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
    <div className="space-y-6 pb-20 text-slate-800 dark:text-slate-200 font-sans transition-colors duration-300">
      
      {/* 🚀 CABECERA PRINCIPAL MODERNA */}
      <div className="bg-white dark:bg-slate-800 p-6 rounded-3xl shadow-xl border border-slate-200 dark:border-slate-700 flex flex-col md:flex-row justify-between items-center gap-4 transition-colors duration-300">
        <div className="space-y-1">
          <div className="flex items-center gap-3">
            <span className="text-2xl p-2 bg-green-700/10 dark:bg-emerald-500/20 rounded-xl text-green-700 dark:text-emerald-400">🏠</span>
            <h2 className="text-xl font-black tracking-tight text-slate-900 dark:text-white uppercase">Configuración de Invernaderos</h2>
          </div>
          <p className="text-xs text-slate-500 dark:text-slate-400 font-medium">Gestión de bloques, cultivos y estados operativos.</p>
        </div>
        
        <div className="flex items-center gap-3 w-full md:w-auto">
          <button onClick={exportarInvernaderosAExcel} className="flex-1 md:flex-none px-4 py-2.5 bg-emerald-600 hover:bg-emerald-700 text-white font-black text-xs uppercase tracking-wider rounded-xl shadow-md transition-all flex items-center justify-center gap-2 cursor-pointer">
            <span>📊</span> Exportar Excel
          </button>
          <button onClick={abrirModalNuevo} className="flex-1 md:flex-none px-5 py-2.5 bg-green-700 hover:bg-green-800 text-white font-black text-xs uppercase tracking-wider rounded-xl shadow-md transition-all flex items-center justify-center gap-2 cursor-pointer">
            <span>+</span> Registrar Invernadero
          </button>
        </div>
      </div>

      {/* 🔍 BARRA DE BÚSQUEDA Y PESTAÑAS */}
      <div className="bg-white dark:bg-slate-800 p-4 rounded-2xl shadow-md border border-slate-200 dark:border-slate-700 flex flex-col md:flex-row justify-between items-center gap-4 transition-colors duration-300">
        <div className="flex items-center gap-2 overflow-x-auto w-full md:w-auto pb-1 md:pb-0">
          <button 
            onClick={() => setTabTabla('activos')} 
            className={`px-4 py-2.5 rounded-xl text-[10px] uppercase font-black transition-all whitespace-nowrap cursor-pointer ${
              tabTabla === 'activos' ? 'bg-green-700 text-white shadow-md' : 'bg-slate-100 dark:bg-slate-700 text-slate-500 dark:text-slate-300 hover:bg-slate-200 dark:hover:bg-slate-600'
            }`}
          >
            🌱 Operativos ({listaActivos.length})
          </button>
          <button 
            onClick={() => setTabTabla('archivados')} 
            className={`px-4 py-2.5 rounded-xl text-[10px] uppercase font-black transition-all whitespace-nowrap cursor-pointer ${
              tabTabla === 'archivados' ? 'bg-red-600 text-white shadow-md' : 'bg-slate-100 dark:bg-slate-700 text-slate-500 dark:text-slate-300 hover:bg-slate-200 dark:hover:bg-slate-600'
            }`}
          >
            📁 Archivados ({listaArchivados.length})
          </button>
        </div>
        
        <div className="relative w-full md:w-96">
          <span className="absolute inset-y-0 left-0 flex items-center pl-3 text-slate-400">🔍</span>
          <input 
            type="text" 
            placeholder="Buscar bloque, cultivo..." 
            className="w-full bg-slate-50 dark:bg-slate-700/50 border border-slate-200 dark:border-slate-600 text-xs text-slate-800 dark:text-white rounded-xl pl-10 pr-4 py-2.5 outline-none focus:border-green-700 font-bold placeholder-slate-400"
            value={busqueda}
            onChange={e => setBusqueda(e.target.value)}
          />
        </div>
      </div>

      {/* 📊 TABLA DE INVERNADEROS (ANCHO COMPLETO) */}
      <div className="bg-white dark:bg-slate-800 rounded-3xl shadow-xl overflow-hidden border border-slate-200 dark:border-slate-700 transition-colors duration-300">
        <div className="overflow-x-auto max-h-[600px]">
          <table className="w-full text-left text-xs border-collapse">
            <thead>
              <tr className="bg-slate-100 dark:bg-slate-700/80 text-slate-700 dark:text-slate-300 uppercase font-black text-[10px] tracking-wider border-b border-slate-200 dark:border-slate-700 sticky top-0">
                <th className="p-4">Invernadero / Bloque</th>
                <th className="p-4">Cultivo / Clasificación</th>
                <th className="p-4 text-center">Estado</th>
                <th className="p-4">Siembra / Est. Cosecha</th>
                <th className="p-4 text-center">Acciones</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100 dark:divide-slate-700/60 font-bold text-slate-700 dark:text-slate-300">
              {listaFiltradaTabla.length === 0 ? (
                <tr>
                  <td colSpan="5" className="p-8 text-center text-slate-400 dark:text-slate-500 font-bold italic">
                    {tabTabla === 'activos' ? 'No hay invernaderos activos en producción.' : 'No hay invernaderos archivados en el histórico.'}
                  </td>
                </tr>
              ) : (
                listaFiltradaTabla.map((item, index) => {
                  const esInactivo = item.activo === false;
                  const estVisual = String(item.estado || '').toUpperCase();

                  return (
                    <tr key={item.id} className={`${index % 2 === 0 ? 'bg-white dark:bg-slate-800' : 'bg-slate-50/50 dark:bg-slate-800/60'} hover:bg-sky-50/50 dark:hover:bg-slate-700/50 transition-colors border-l-8 ${esInactivo ? 'border-red-600' : 'border-green-700'}`}>
                      
                      <td className="p-4 font-black text-slate-900 dark:text-white whitespace-nowrap">
                        <p className="uppercase text-sm">{item.nombre}</p>
                        {esInactivo && <span className="text-[9px] text-red-500 dark:text-red-400 font-bold italic block mt-0.5">📁 Registros Conservados</span>}
                      </td>
                      
                      <td className="p-4 font-bold text-slate-700 dark:text-slate-300 uppercase">
                        <p className="text-slate-900 dark:text-white">{item.cultivo || item.cultivo_principal || 'S/C'} <span className="text-slate-400 font-normal">/</span> {item.variedad || 'S/C'}</p>
                        <p className="text-[10px] text-slate-500 dark:text-slate-400 font-bold italic mt-0.5">📐 Área: {item.largo || 0}m × {item.ancho || 0}m ({(parseFloat(item.largo || 0) * parseFloat(item.ancho || 0)).toLocaleString()} m²)</p>
                      </td>
                      
                      <td className="p-4 text-center whitespace-nowrap">
                        {esInactivo ? (
                          <span className="px-2.5 py-1 rounded-md text-[9px] font-black uppercase shadow-sm bg-red-100 dark:bg-red-950/80 text-red-700 dark:text-red-400 border border-red-200 dark:border-red-800">
                            ARCHIVADO
                          </span>
                        ) : (
                          <span className={`px-2.5 py-1 rounded-md text-[9px] font-black uppercase shadow-sm ${
                            estVisual === 'ACTIVO' ? 'bg-green-100 dark:bg-emerald-950/80 text-green-700 dark:text-emerald-400 border border-green-200 dark:border-emerald-800' : 
                            estVisual === 'EN_COSECHA' ? 'bg-blue-100 dark:bg-blue-950/80 text-blue-700 dark:text-blue-400 border border-blue-200 dark:border-blue-800' :
                            estVisual === 'EN_PREPARACION' ? 'bg-amber-100 dark:bg-amber-950/80 text-amber-700 dark:text-amber-400 border border-amber-200 dark:border-amber-800' : 
                            estVisual === 'EN_DESCANSO' || estVisual === 'INACTIVO' ? 'bg-purple-100 dark:bg-purple-950/80 text-purple-700 dark:text-purple-400 border border-purple-200 dark:border-purple-800' :
                            'bg-slate-200 dark:bg-slate-700 text-slate-700 dark:text-slate-300 border border-slate-300 dark:border-slate-600'
                          }`}>
                            {estVisual === 'EN_DESCANSO' || (estVisual === 'INACTIVO' && !esInactivo) ? 'EN DESCANSO' : estVisual.replace('_', ' ')}
                          </span>
                        )}
                      </td>

                      <td className="p-4 font-bold text-slate-600 dark:text-slate-400 whitespace-nowrap">
                        <p className="text-[11px]">🌱 <span className="text-slate-400 font-black">S:</span> {item.fecha_siembra || 'N/R'}</p>
                        <p className="text-[11px] mt-1">🚜 <span className="text-slate-400 font-black">C:</span> {item.fecha_cosecha_est || 'N/R'}</p>
                      </td>

                      <td className="p-4 whitespace-nowrap">
                        <div className="flex gap-1.5 justify-center">
                          <button
                            type="button"
                            onClick={() => prepararEdicionInvernadero(item)}
                            className="px-2.5 py-1.5 bg-slate-700 dark:bg-slate-600 text-white hover:bg-slate-900 dark:hover:bg-slate-500 rounded-lg shadow-sm border border-slate-800 dark:border-slate-500 transition-colors text-[10px] font-black flex items-center gap-1 cursor-pointer"
                            title="Editar Invernadero"
                          >
                            <span>✏️</span> EDITAR
                          </button>
                          
                          {!esInactivo ? (
                            <button
                              type="button"
                              onClick={() => handleInactivarLogico(item.id, item.nombre)}
                              className="p-1.5 bg-red-100 dark:bg-red-950/60 text-red-700 dark:text-red-400 hover:bg-red-700 hover:text-white rounded-lg border border-red-200 dark:border-red-900 transition-colors cursor-pointer"
                              title="Archivar Invernadero"
                            >
                              🚫
                            </button>
                          ) : (
                            <button
                              type="button"
                              onClick={() => handleReactivar(item.id, item.nombre)}
                              className="p-1.5 bg-emerald-100 dark:bg-emerald-950/60 text-emerald-700 dark:text-emerald-400 hover:bg-emerald-600 hover:text-white rounded-lg border border-emerald-200 dark:border-emerald-900 transition-colors cursor-pointer"
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

      {/* 🧊 MODAL FLOTANTE PARA NUEVO / EDITAR INVERNADERO */}
      {modalAbierto && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 backdrop-blur-sm p-4 overflow-y-auto">
          <div className="bg-white dark:bg-slate-800 border border-slate-200 dark:border-slate-700 w-full max-w-2xl rounded-3xl shadow-2xl overflow-hidden animate-in fade-in zoom-in-95 duration-200 my-8 text-slate-800 dark:text-slate-200">
            
            <div className="p-5 bg-slate-900 text-white border-b border-slate-800 flex justify-between items-center">
              <div className="flex items-center gap-3">
                <span className="text-xl">🏠</span>
                <div>
                  <h3 className="text-base font-black uppercase tracking-tight">
                    {invForm.id_editando ? "Editar Invernadero" : "Nuevo Invernadero"}
                  </h3>
                  <p className="text-[10px] text-slate-400 font-medium">Configuración de área, cultivo y tiempos operativos.</p>
                </div>
              </div>
              <button onClick={() => setModalAbierto(false)} className="text-slate-400 hover:text-white text-lg font-black px-3 py-1 rounded-xl bg-slate-800 cursor-pointer">✕</button>
            </div>

            <div className="p-6 space-y-4 max-h-[80vh] overflow-y-auto">
              
              <div>
                <label className="text-[10px] font-black text-slate-500 dark:text-slate-400 uppercase tracking-wider">Nombre / Identificador *</label>
                <input 
                  type="text" 
                  className="w-full bg-slate-50 dark:bg-slate-900 border border-slate-200 dark:border-slate-700 text-slate-800 dark:text-white p-2.5 rounded-xl font-bold text-sm uppercase outline-none focus:border-green-700 mt-1 placeholder-slate-400" 
                  value={invForm.nombre || ''} 
                  onChange={e => setInvForm({...invForm, nombre: e.target.value})} 
                  placeholder="Ej: INVERNADERO 1" 
                />
              </div>

              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                <div>
                  <label className="text-[10px] font-black text-slate-500 dark:text-slate-400 uppercase tracking-wider">Tipo de Cultivo</label>
                  <select 
                    className="w-full bg-slate-50 dark:bg-slate-900 border border-slate-200 dark:border-slate-700 text-slate-800 dark:text-white p-2.5 rounded-xl font-bold text-xs uppercase outline-none focus:border-green-700 mt-1" 
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
                  <label className="text-[10px] font-black text-slate-500 dark:text-slate-400 uppercase tracking-wider">Clasificación / Variedad</label>
                  <select 
                    className="w-full bg-slate-50 dark:bg-slate-900 border border-slate-200 dark:border-slate-700 text-slate-800 dark:text-white p-2.5 rounded-xl font-bold text-xs uppercase outline-none focus:border-green-700 mt-1" 
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

              <div className="bg-slate-50 dark:bg-slate-900/60 p-4 rounded-2xl border border-slate-200 dark:border-slate-700 space-y-3 shadow-inner">
                <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                  <div>
                    <label className="text-[10px] font-black text-slate-500 dark:text-slate-400 uppercase tracking-wider">Largo (Metros)</label>
                    <input 
                      type="number" 
                      className="w-full bg-white dark:bg-slate-800 border border-slate-200 dark:border-slate-700 text-slate-800 dark:text-white p-2.5 rounded-xl font-bold text-sm outline-none focus:border-green-700 mt-1 placeholder-slate-400" 
                      value={invForm.largo || ''} 
                      onChange={e => setInvForm({...invForm, largo: e.target.value})} 
                      placeholder="0.00" 
                    />
                  </div>
                  <div>
                    <label className="text-[10px] font-black text-slate-500 dark:text-slate-400 uppercase tracking-wider">Ancho (Metros)</label>
                    <input 
                      type="number" 
                      className="w-full bg-white dark:bg-slate-800 border border-slate-200 dark:border-slate-700 text-slate-800 dark:text-white p-2.5 rounded-xl font-bold text-sm outline-none focus:border-green-700 mt-1 placeholder-slate-400" 
                      value={invForm.ancho || ''} 
                      onChange={e => setInvForm({...invForm, ancho: e.target.value})} 
                      placeholder="0.00" 
                    />
                  </div>
                </div>
              </div>

              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                <div>
                  <label className="text-[10px] font-black text-slate-500 dark:text-slate-400 uppercase tracking-wider">Fecha Siembra</label>
                  <input 
                    type="date" 
                    className="w-full bg-slate-50 dark:bg-slate-900 border border-slate-200 dark:border-slate-700 text-slate-800 dark:text-white p-2.5 rounded-xl font-bold text-xs outline-none focus:border-green-700 mt-1" 
                    value={invForm.siembra || ''} 
                    onChange={e => setInvForm({...invForm, siembra: e.target.value})} 
                  />
                </div>
                <div>
                  <label className="text-[10px] font-black text-slate-500 dark:text-slate-400 uppercase tracking-wider">Est. Cosecha</label>
                  <input 
                    type="date" 
                    className="w-full bg-slate-50 dark:bg-slate-900 border border-slate-200 dark:border-slate-700 text-slate-800 dark:text-white p-2.5 rounded-xl font-bold text-xs outline-none focus:border-green-700 mt-1" 
                    value={invForm.cosecha || ''} 
                    onChange={e => setInvForm({...invForm, cosecha: e.target.value})} 
                  />
                </div>
              </div>

              <div>
                <label className="text-[10px] font-black text-slate-500 dark:text-slate-400 uppercase tracking-wider">Estado Actual del Bloque</label>
                <select 
                  className="w-full bg-slate-50 dark:bg-slate-900 border border-slate-200 dark:border-slate-700 text-slate-800 dark:text-white p-2.5 rounded-xl font-black text-xs outline-none focus:border-green-700 uppercase mt-1" 
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
                <label className="text-[10px] font-black text-slate-500 dark:text-slate-400 uppercase tracking-wider">Notas / Descripción</label>
                <textarea 
                  className="w-full bg-slate-50 dark:bg-slate-900 border border-slate-200 dark:border-slate-700 text-slate-800 dark:text-white p-2.5 rounded-xl font-bold h-16 text-xs uppercase outline-none focus:border-green-700 mt-1 resize-none placeholder-slate-400" 
                  value={invForm.descripcion || ''} 
                  onChange={e => setInvForm({...invForm, descripcion: e.target.value})} 
                  placeholder="Ej: Suelo tratado con micorrizas" 
                />
              </div>

              <div className="flex justify-end gap-3 pt-4 border-t border-slate-200 dark:border-slate-700">
                <button type="button" onClick={() => setModalAbierto(false)} className="px-5 py-2.5 bg-slate-200 dark:bg-slate-700 hover:bg-slate-300 dark:hover:bg-slate-600 text-slate-700 dark:text-slate-200 font-black text-xs uppercase tracking-wider rounded-xl transition-colors cursor-pointer">
                  Cancelar
                </button>
                <button type="button" onClick={handleSave} className="px-6 py-2.5 bg-green-700 hover:bg-green-800 text-white font-black text-xs uppercase tracking-wider rounded-xl shadow-lg transition-colors cursor-pointer">
                  {invForm.id_editando ? '💾 Guardar Cambios' : '🚀 Crear Invernadero'}
                </button>
              </div>

            </div>
          </div>
        </div>
      )}

    </div>
  );
}