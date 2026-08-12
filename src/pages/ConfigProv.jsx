import React, { useState } from 'react';
import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';

export default function ConfigProv({ provForm, setProvForm, mostrarAlerta, cargarTodo, supabase, lista }) {
  
  const [bancoOtroTexto, setBancoOtroTexto] = useState('');
  const [busqueda, setBusqueda] = useState(''); 
  const [modalAbierto, setModalAbierto] = useState(false); 

  const bancosSoportados = ["Efectivo", "Bre-B (Pago Inmediato)", "Bancolombia Ahorros", "Bancolombia Corriente", "Nequi", "Daviplata", "Banco de Bogotá", "Colpatria", "Davivienda", "Otro"];

  const handleSave = async () => {
    if (!provForm.nombre || !provForm.nit) {
      mostrarAlerta("Nombre y NIT/CC son obligatorios", "error");
      return;
    }

    let bancoFinal = provForm.banco || "Efectivo";
    if (provForm.banco === 'Otro') {
      bancoFinal = bancoOtroTexto.trim() !== '' ? bancoOtroTexto.toUpperCase().trim() : 'Otro';
    }

    const payload = {
      nombre: provForm.nombre.toUpperCase().trim(),
      nit_cc: provForm.nit.toString().trim(),
      telefono: provForm.tel ? provForm.tel.toString().trim() : null,
      email: provForm.email ? provForm.email.toLowerCase().trim() : null,
      direccion: provForm.dir ? provForm.dir.toUpperCase().trim() : null,
      ciudad: provForm.ciudad ? provForm.ciudad.toUpperCase().trim() : null,
      nota: provForm.nota ? provForm.nota.trim() : null,
      banco: bancoFinal, 
      numero_cuenta: provForm.numero_cuenta ? provForm.numero_cuenta.trim() : null 
    };

    try {
      if (provForm.id_editando) {
        const { error: updateError } = await supabase
          .from('proveedores')
          .update(payload)
          .eq('id', provForm.id_editando);

        if (updateError) throw updateError;
        mostrarAlerta("Proveedor actualizado con éxito", "exito");
      } else {
        const { error: insertError } = await supabase
          .from('proveedores')
          .insert([payload]);

        if (insertError) throw insertError;
        mostrarAlerta("Proveedor guardado correctamente", "exito");
      }

      limpiarFormulario();
      setModalAbierto(false);
      await cargarTodo();
    } catch (error) {
      console.error("Error en la operación de proveedores:", error);
      mostrarAlerta("Error en base de datos: " + (error.message || error), "error");
    }
  };

  const prepararEdicion = (item) => {
    const esOpPredefinida = bancosSoportados.includes(item.banco);
    
    setProvForm({
      id_editando: item.id,
      nombre: item.nombre || '',
      nit: item.nit_cc || '',
      tel: item.telefono || '',
      email: item.email || '',
      dir: item.direccion || '',
      ciudad: item.ciudad || '',
      nota: item.nota || '',
      banco: esOpPredefinida ? (item.banco || 'Efectivo') : 'Otro', 
      numero_cuenta: item.numero_cuenta || '' 
    });

    if (!esOpPredefinida) {
      setBancoOtroTexto(item.banco || '');
    } else {
      setBancoOtroTexto('');
    }

    setModalAbierto(true);
  };

  const eliminarProveedor = async (id, nombre) => {
    if (window.confirm(`¿Estás seguro de eliminar al proveedor "${nombre}"? Esta acción no se puede deshacer.`)) {
      try {
        const { error } = await supabase.from('proveedores').delete().eq('id', id);
        if (error) throw error;
        
        mostrarAlerta("Proveedor eliminado definitivamente", "exito");
        await cargarTodo();
      } catch (err) {
        console.error("Error al eliminar:", err);
        mostrarAlerta("No se puede eliminar: El proveedor tiene egresos o registros asociados en el historial.", "error");
      }
    }
  };

  const limpiarFormulario = () => {
    setProvForm({ 
      id_editando: null, 
      nombre: '', 
      nit: '', 
      tel: '', 
      email: '',
      dir: '', 
      ciudad: '', 
      nota: '',
      banco: 'Efectivo',
      numero_cuenta: '' 
    });
    setBancoOtroTexto('');
  };

  const abrirModalNuevo = () => {
    limpiarFormulario();
    setModalAbierto(true);
  };

  const exportarProveedoresAExcel = async () => {
    if (!lista || lista.length === 0) {
      mostrarAlerta("No hay proveedores para exportar", "error");
      return;
    }

    try {
      const workbook = new ExcelJS.Workbook();
      const sheet = workbook.addWorksheet('Directorio Proveedores');

      sheet.columns = [
        { header: 'RAZÓN SOCIAL / NOMBRE', key: 'nombre', width: 32 },
        { header: 'NIT / CÉDULA', key: 'nit', width: 18 },
        { header: 'FORMA / BANCO PAGO', key: 'banco', width: 22 },
        { header: 'N° CUENTA / LLAVE', key: 'cuenta', width: 22 },
        { header: 'TELÉFONO', key: 'telefono', width: 18 },
        { header: 'CORREO ELECTRÓNICO', key: 'email', width: 28 },
        { header: 'CIUDAD', key: 'ciudad', width: 18 },
        { header: 'DIRECCIÓN FÍSICA', key: 'direccion', width: 28 },
        { header: 'OBSERVACIONES / NOTAS', key: 'nota', width: 35 }
      ];

      lista.forEach(p => {
        sheet.addRow({
          nombre: (p.nombre || '').toUpperCase(),
          nit: p.nit_cc || '',
          banco: (p.banco || 'Efectivo').toUpperCase(),
          cuenta: p.numero_cuenta || '---',
          telefono: p.telefono || 'N/R',
          email: p.email ? p.email.toLowerCase() : 'N/R',
          ciudad: (p.ciudad || 'N/R').toUpperCase(),
          direccion: (p.direccion || 'N/R').toUpperCase(),
          nota: p.nota || ''
        });
      });

      const headerRow = sheet.getRow(1);
      headerRow.height = 24;
      headerRow.eachCell(cell => {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF117097' } };
        cell.font = { name: 'Arial', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
      });

      sheet.eachRow((row, rowNumber) => {
        if (rowNumber === 1) return;
        row.height = 20;
        const esCebra = rowNumber % 2 === 0;
        row.eachCell((cell, colNumber) => {
          cell.font = { name: 'Arial', size: 9 };
          if (esCebra) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEBF5FB' } };

          if ([2, 3, 4, 5, 7].includes(colNumber)) cell.alignment = { vertical: 'middle', horizontal: 'center' };
          else cell.alignment = { vertical: 'middle', horizontal: 'left' };
        });
      });

      sheet.autoFilter = { from: { row: 1, column: 1 }, to: { row: sheet.rowCount, column: sheet.columnCount } };

      const buffer = await workbook.xlsx.writeBuffer();
      const fechaHoy = new Date().toISOString().split('T')[0];
      const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      saveAs(blob, `BASE_DATOS_PROVEEDORES_${fechaHoy}.xlsx`);

      mostrarAlerta("Base de datos de proveedores exportada a Excel", "exito");
    } catch (err) {
      console.error("Error al exportar Excel:", err);
      mostrarAlerta("Error al generar el archivo Excel", "error");
    }
  };

  const proveedoresFiltrados = (lista || []).filter(p => {
    const query = busqueda.toLowerCase();
    return (
      (p.nombre || '').toLowerCase().includes(query) ||
      (p.nit_cc || '').toLowerCase().includes(query) ||
      (p.ciudad || '').toLowerCase().includes(query) ||
      (p.banco || '').toLowerCase().includes(query)
    );
  });

  return (
    <div className="space-y-6 pb-20 text-slate-800 dark:text-slate-200 font-sans transition-colors duration-300">
      
      {/* 🚀 CABECERA PRINCIPAL MODERNA */}
      <div className="bg-white dark:bg-slate-800 p-6 rounded-3xl shadow-xl border border-slate-200 dark:border-slate-700 flex flex-col md:flex-row justify-between items-center gap-4 transition-colors duration-300">
        <div className="space-y-1">
          <div className="flex items-center gap-3">
            <span className="text-2xl p-2 bg-[#117097]/10 dark:bg-sky-500/20 rounded-xl text-[#117097] dark:text-sky-400">🚛</span>
            <h2 className="text-xl font-black tracking-tight text-slate-900 dark:text-white uppercase">Directorio de Proveedores</h2>
          </div>
          <p className="text-xs text-slate-500 dark:text-slate-400 font-medium">Gestión de base de datos logística y de abastecimiento.</p>
        </div>
        
        <div className="flex items-center gap-3 w-full md:w-auto">
          <button onClick={exportarProveedoresAExcel} className="flex-1 md:flex-none px-4 py-2.5 bg-emerald-600 hover:bg-emerald-700 text-white font-black text-xs uppercase tracking-wider rounded-xl shadow-md transition-all flex items-center justify-center gap-2 cursor-pointer">
            <span>📊</span> Exportar Excel
          </button>
          <button onClick={abrirModalNuevo} className="flex-1 md:flex-none px-5 py-2.5 bg-[#117097] hover:bg-[#0a4c68] text-white font-black text-xs uppercase tracking-wider rounded-xl shadow-md transition-all flex items-center justify-center gap-2 cursor-pointer">
            <span>+</span> Nuevo Proveedor
          </button>
        </div>
      </div>

      {/* 🔍 BARRA DE BÚSQUEDA Y TOTALES */}
      <div className="bg-white dark:bg-slate-800 p-4 rounded-2xl shadow-md border border-slate-200 dark:border-slate-700 flex flex-col md:flex-row justify-between items-center gap-4 transition-colors duration-300">
        <div className="flex items-center gap-3">
          <span className="px-4 py-2 bg-slate-100 dark:bg-slate-700 text-slate-700 dark:text-slate-300 font-black text-[10px] uppercase rounded-xl border border-slate-200 dark:border-slate-600">
            Total Proveedores: {proveedoresFiltrados.length}
          </span>
          <span className="px-3 py-1.5 bg-[#117097]/10 dark:bg-sky-950/60 text-[#117097] dark:text-sky-400 font-black text-[10px] uppercase rounded-lg border border-[#117097]/20 dark:border-sky-800">
            Abastecimiento
          </span>
        </div>
        
        <div className="relative w-full md:w-96">
          <span className="absolute inset-y-0 left-0 flex items-center pl-3 text-slate-400">🔍</span>
          <input 
            type="text" 
            placeholder="Buscar por nombre, NIT, ciudad, banco..." 
            className="w-full bg-slate-50 dark:bg-slate-700/50 border border-slate-200 dark:border-slate-600 text-xs text-slate-800 dark:text-white rounded-xl pl-10 pr-4 py-2.5 outline-none focus:border-[#117097] font-bold placeholder-slate-400"
            value={busqueda}
            onChange={e => setBusqueda(e.target.value)}
          />
        </div>
      </div>

      {/* 📊 TABLA DE PROVEEDORES (ANCHO COMPLETO) */}
      <div className="bg-white dark:bg-slate-800 rounded-3xl shadow-xl overflow-hidden border border-slate-200 dark:border-slate-700 transition-colors duration-300">
        <div className="overflow-x-auto max-h-[600px]">
          <table className="w-full text-left text-xs border-collapse">
            <thead>
              <tr className="bg-slate-100 dark:bg-slate-700/80 text-slate-700 dark:text-slate-300 uppercase font-black text-[10px] tracking-wider border-b border-slate-200 dark:border-slate-700 sticky top-0">
                <th className="p-4">Nombre / Razón Social</th>
                <th className="p-4">NIT / CC</th>
                <th className="p-4">Datos de Pago</th>
                <th className="p-4 text-center">Contacto / Ubicación</th>
                <th className="p-4 text-center">Acciones</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100 dark:divide-slate-700/60 font-bold text-slate-700 dark:text-slate-300">
              {proveedoresFiltrados.length === 0 ? (
                <tr>
                  <td colSpan="5" className="p-8 text-center text-slate-400 dark:text-slate-500 font-bold italic">
                    No hay proveedores registrados o coincidentes con la búsqueda.
                  </td>
                </tr>
              ) : (
                proveedoresFiltrados.map((item, index) => (
                  <tr key={item.id} className={`${index % 2 === 0 ? 'bg-white dark:bg-slate-800' : 'bg-slate-50/50 dark:bg-slate-800/60'} hover:bg-sky-50/50 dark:hover:bg-slate-700/50 transition-colors border-l-8 border-[#117097] dark:border-sky-600`}>
                    
                    <td className="p-4 font-black text-slate-900 dark:text-white">
                      <p className="uppercase text-sm leading-snug">{item.nombre}</p>
                      {item.nota && <p className="text-[10px] text-slate-400 dark:text-slate-400 font-medium normal-case mt-1 line-clamp-2">📝 {item.nota}</p>}
                    </td>
                    
                    <td className="p-4 font-bold text-slate-700 dark:text-slate-300 whitespace-nowrap">
                      {item.nit_cc}
                    </td>
                    
                    {/* DATOS DE PAGO */}
                    <td className="p-4 font-bold">
                      <span className={`inline-block px-2 py-0.5 rounded text-[8px] uppercase tracking-wider font-black shadow-sm ${
                        item.banco === 'Efectivo' 
                          ? 'bg-emerald-100 dark:bg-emerald-950/80 text-emerald-800 dark:text-emerald-400' 
                          : item.banco?.toLowerCase().includes('bre-b')
                          ? 'bg-cyan-100 dark:bg-cyan-950/80 text-cyan-800 dark:text-cyan-400 border border-cyan-200 dark:border-cyan-800' 
                          : item.banco?.toLowerCase().includes('nequi') 
                          ? 'bg-purple-100 dark:bg-purple-950/80 text-purple-800 dark:text-purple-400' 
                          : item.banco?.toLowerCase().includes('daviplata')
                          ? 'bg-red-100 dark:bg-red-950/80 text-red-800 dark:text-red-400'
                          : 'bg-blue-100 dark:bg-blue-950/80 text-blue-800 dark:text-blue-400'
                      }`}>
                        {item.banco?.toLowerCase().includes('bre-b') ? '⚡ Bre-B' : `🏦 ${item.banco || 'Efectivo'}`}
                      </span>
                      {item.numero_cuenta && (
                        <p className="text-[10px] text-slate-600 dark:text-slate-300 mt-1.5 font-black bg-slate-100 dark:bg-slate-700 border border-slate-200 dark:border-slate-600 px-1.5 py-0.5 rounded inline-block">
                          {item.banco?.toLowerCase().includes('bre-b') ? `Llave: ${item.numero_cuenta}` : `#${item.numero_cuenta}`}
                        </p>
                      )}
                    </td>

                    {/* CONTACTO Y UBICACIÓN */}
                    <td className="p-4 font-bold text-slate-600 dark:text-slate-300 space-y-1">
                      <p className="flex items-center justify-center gap-1.5 whitespace-nowrap text-xs"><span>📞</span> {item.telefono || 'N/R'}</p>
                      {item.email && <p className="text-[10px] text-[#117097] dark:text-sky-400 font-bold lowercase text-center truncate max-w-[150px] mx-auto" title={item.email}>✉️ {item.email}</p>}
                      <p className="uppercase text-slate-500 dark:text-slate-400 text-[9px] text-center bg-slate-100 dark:bg-slate-700 px-2 py-0.5 rounded border border-slate-200 dark:border-slate-600 w-fit mx-auto mt-1">📍 {item.ciudad || 'N/R'}</p>
                    </td>

                    <td className="p-4 whitespace-nowrap">
                      <div className="flex gap-1.5 justify-center">
                        <button
                          type="button"
                          onClick={() => prepararEdicion(item)}
                          className="px-2.5 py-1.5 bg-slate-700 dark:bg-slate-600 text-white hover:bg-slate-900 dark:hover:bg-slate-500 rounded-lg shadow-sm border border-slate-800 dark:border-slate-500 transition-colors text-[10px] font-black flex items-center gap-1 cursor-pointer"
                          title="Editar Proveedor"
                        >
                          <span>✏️</span> EDITAR
                        </button>
                        
                        <button
                          type="button"
                          onClick={() => eliminarProveedor(item.id, item.nombre)}
                          className="p-1.5 bg-red-100 dark:bg-red-950/60 text-red-700 dark:text-red-400 hover:bg-red-700 hover:text-white rounded-lg border border-red-200 dark:border-red-900 transition-colors cursor-pointer"
                          title="Eliminar Proveedor"
                        >
                          🗑️
                        </button>
                      </div>
                    </td>

                  </tr>
                ))
              )}
            </tbody>
          </table>
        </div>
      </div>

      {/* 🧊 MODAL FLOTANTE PARA NUEVO / EDITAR PROVEEDOR */}
      {modalAbierto && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 backdrop-blur-sm p-4 overflow-y-auto">
          <div className="bg-white dark:bg-slate-800 border border-slate-200 dark:border-slate-700 w-full max-w-2xl rounded-3xl shadow-2xl overflow-hidden animate-in fade-in zoom-in-95 duration-200 my-8 text-slate-800 dark:text-slate-200">
            
            <div className="p-5 bg-slate-900 text-white border-b border-slate-800 flex justify-between items-center">
              <div className="flex items-center gap-3">
                <span className="text-xl">🚛</span>
                <div>
                  <h3 className="text-base font-black uppercase tracking-tight">
                    {provForm.id_editando ? "Editar Proveedor" : "Nuevo Proveedor"}
                  </h3>
                  <p className="text-[10px] text-slate-400 font-medium">Información logística y datos de pago.</p>
                </div>
              </div>
              <button onClick={() => setModalAbierto(false)} className="text-slate-400 hover:text-white text-lg font-black px-3 py-1 rounded-xl bg-slate-800 cursor-pointer">✕</button>
            </div>

            <div className="p-6 space-y-4 max-h-[80vh] overflow-y-auto">
              
              <div>
                <label className="text-[10px] font-black text-slate-500 dark:text-slate-400 uppercase tracking-wider">Razón Social / Nombre *</label>
                <input 
                  type="text" 
                  className="w-full bg-slate-50 dark:bg-slate-900 border border-slate-200 dark:border-slate-700 text-slate-800 dark:text-white p-2.5 rounded-xl font-bold text-sm uppercase outline-none focus:border-[#117097] mt-1 placeholder-slate-400" 
                  value={provForm.nombre || ''} 
                  onChange={e => setProvForm({...provForm, nombre: e.target.value})} 
                  placeholder="Ej: AGROINSUMOS SAS" 
                />
              </div>

              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                <div>
                  <label className="text-[10px] font-black text-slate-500 dark:text-slate-400 uppercase tracking-wider">NIT / Cédula *</label>
                  <input 
                    type="text" 
                    className="w-full bg-slate-50 dark:bg-slate-900 border border-slate-200 dark:border-slate-700 text-slate-800 dark:text-white p-2.5 rounded-xl font-bold text-sm uppercase outline-none focus:border-[#117097] mt-1 placeholder-slate-400" 
                    value={provForm.nit || ''} 
                    onChange={e => setProvForm({...provForm, nit: e.target.value})} 
                    placeholder="800.000.000-1" 
                  />
                </div>
                <div>
                  <label className="text-[10px] font-black text-slate-500 dark:text-slate-400 uppercase tracking-wider">Correo Electrónico (E-Mail)</label>
                  <input 
                    type="email" 
                    className="w-full bg-slate-50 dark:bg-slate-900 border border-slate-200 dark:border-slate-700 text-slate-800 dark:text-white p-2.5 rounded-xl font-bold text-xs lowercase outline-none focus:border-[#117097] mt-1 placeholder-slate-400" 
                    value={provForm.email || ''} 
                    onChange={e => setProvForm({...provForm, email: e.target.value})} 
                    placeholder="contacto@proveedor.com" 
                  />
                </div>
              </div>

              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                <div>
                  <label className="text-[10px] font-black text-slate-500 dark:text-slate-400 uppercase tracking-wider">Teléfono / Celular</label>
                  <input 
                    type="text" 
                    className="w-full bg-slate-50 dark:bg-slate-900 border border-slate-200 dark:border-slate-700 text-slate-800 dark:text-white p-2.5 rounded-xl font-bold text-xs uppercase outline-none focus:border-[#117097] mt-1 placeholder-slate-400" 
                    value={provForm.tel || ''} 
                    onChange={e => setProvForm({...provForm, tel: e.target.value})} 
                    placeholder="315..." 
                  />
                </div>
                <div>
                  <label className="text-[10px] font-black text-slate-500 dark:text-slate-400 uppercase tracking-wider">Ciudad</label>
                  <input 
                    type="text" 
                    className="w-full bg-slate-50 dark:bg-slate-900 border border-slate-200 dark:border-slate-700 text-slate-800 dark:text-white p-2.5 rounded-xl font-bold text-xs uppercase outline-none focus:border-[#117097] mt-1 placeholder-slate-400" 
                    value={provForm.ciudad || ''} 
                    onChange={e => setProvForm({...provForm, ciudad: e.target.value})} 
                    placeholder="Chía" 
                  />
                </div>
              </div>

              <div>
                <label className="text-[10px] font-black text-slate-500 dark:text-slate-400 uppercase tracking-wider">Dirección Físíca</label>
                <input 
                  type="text" 
                  className="w-full bg-slate-50 dark:bg-slate-900 border border-slate-200 dark:border-slate-700 text-slate-800 dark:text-white p-2.5 rounded-xl font-bold text-xs uppercase outline-none focus:border-[#117097] mt-1 placeholder-slate-400" 
                  value={provForm.dir || ''} 
                  onChange={e => setProvForm({...provForm, dir: e.target.value})} 
                  placeholder="Zona Industrial..." 
                />
              </div>

              {/* 💳 SECCIÓN: MEDIOS DE PAGO CON ENTRADA MANUAL DINÁMICA */}
              <div className="bg-slate-50 dark:bg-slate-900/60 p-4 rounded-2xl border border-slate-200 dark:border-slate-700 space-y-3 shadow-inner">
                <p className="text-[10px] font-black text-[#117097] dark:text-sky-400 uppercase tracking-wider border-b border-slate-200 dark:border-slate-700 pb-1.5">Información de Pago / Transferencia</p>
                
                <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                  <div>
                    <label className="text-[9px] font-black text-slate-500 dark:text-slate-400 uppercase">Forma / Banco</label>
                    <select 
                      className="w-full bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-700 text-slate-800 dark:text-white p-2.5 rounded-xl outline-none text-xs font-bold focus:border-[#117097] mt-1" 
                      value={provForm.banco || 'Efectivo'} 
                      onChange={e => setProvForm({...provForm, banco: e.target.value})}
                    >
                      {bancosSoportados.map(b => <option key={b} value={b}>{b}</option>)}
                    </select>
                  </div>
                  <div>
                    <label className="text-[9px] font-black text-slate-500 dark:text-slate-400 uppercase">N° Cuenta / Llave Bre-B</label>
                    <input 
                      type="text" 
                      className="w-full bg-white dark:bg-slate-900 border border-slate-200 dark:border-slate-700 p-2.5 rounded-xl outline-none text-xs font-black text-slate-700 dark:text-white focus:border-[#117097] mt-1 placeholder-slate-400" 
                      value={provForm.numero_cuenta || ''} 
                      onChange={e => setProvForm({...provForm, numero_cuenta: e.target.value})} 
                      placeholder="Celular, cédula o cuenta..." 
                    />
                  </div>
                </div>

                {provForm.banco === 'Otro' && (
                  <div className="pt-2 animate-in fade-in duration-300">
                    <label className="text-[9px] font-black text-amber-600 dark:text-amber-400 uppercase tracking-wider">Especifique el Banco / Entidad</label>
                    <input 
                      type="text" 
                      className="w-full bg-amber-50 dark:bg-amber-950/40 border border-dashed border-amber-300 dark:border-amber-700 p-2.5 rounded-xl outline-none text-xs font-black uppercase text-amber-700 dark:text-amber-300 mt-1 focus:border-amber-500 placeholder-slate-400" 
                      value={bancoOtroTexto} 
                      onChange={e => setBancoOtroTexto(e.target.value)} 
                      placeholder="Ej: BANCO BBVA / ITAÚ / BANCO AGRARIO..." 
                    />
                  </div>
                )}
              </div>

              <div>
                <label className="text-[10px] font-black text-slate-500 dark:text-slate-400 uppercase tracking-wider">Observaciones / Notas</label>
                <textarea 
                  className="w-full bg-slate-50 dark:bg-slate-900 border border-slate-200 dark:border-slate-700 text-slate-800 dark:text-white p-2.5 rounded-xl font-bold h-20 text-xs outline-none focus:border-[#117097] mt-1 resize-none placeholder-slate-400" 
                  value={provForm.nota || ''} 
                  onChange={e => setProvForm({...provForm, nota: e.target.value})} 
                  placeholder="Detalles de entrega, horarios..." 
                />
              </div>

              <div className="flex justify-end gap-3 pt-4 border-t border-slate-200 dark:border-slate-700">
                <button type="button" onClick={() => setModalAbierto(false)} className="px-5 py-2.5 bg-slate-200 dark:bg-slate-700 hover:bg-slate-300 dark:hover:bg-slate-600 text-slate-700 dark:text-slate-200 font-black text-xs uppercase tracking-wider rounded-xl transition-colors cursor-pointer">
                  Cancelar
                </button>
                <button type="button" onClick={handleSave} className="px-6 py-2.5 bg-[#117097] hover:bg-[#0a4c68] text-white font-black text-xs uppercase tracking-wider rounded-xl shadow-lg transition-colors cursor-pointer">
                  {provForm.id_editando ? '💾 Guardar Cambios' : '🚀 Registrar Proveedor'}
                </button>
              </div>

            </div>
          </div>
        </div>
      )}

    </div>
  );
}