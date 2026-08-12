import React, { useState } from 'react';
import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';

export default function ConfigCli({ cliForm, setCliForm, mostrarAlerta, cargarTodo, supabase, lista }) {
  const [busqueda, setBusqueda] = useState('');
  const [modalAbierto, setModalAbierto] = useState(false);

  const handleSave = async () => {
    if (!cliForm.nombre || !cliForm.nit) {
      mostrarAlerta("Nombre y NIT son obligatorios", "error");
      return;
    }

    const payload = {
      nombre_completo: cliForm.nombre.toUpperCase().trim(),
      nit_cc: cliForm.nit.toString().trim(),
      telefono: cliForm.tel ? cliForm.tel.toString().trim() : null,
      correo: cliForm.email ? cliForm.email.toLowerCase().trim() : null,
      direccion: cliForm.dir ? cliForm.dir.toUpperCase().trim() : null,
      ciudad: cliForm.ciudad ? cliForm.ciudad.toUpperCase().trim() : null,
      nota: cliForm.nota ? cliForm.nota.trim() : null,
      activo: true
    };

    try {
      if (cliForm.id_editando) {
        const { error: updateError } = await supabase
          .from('clientes')
          .update(payload)
          .eq('id', cliForm.id_editando);

        if (updateError) throw updateError;
        mostrarAlerta("Cliente actualizado con éxito", "exito");
      } else {
        const { error: insertError } = await supabase
          .from('clientes')
          .insert([payload]);

        if (insertError) throw insertError;
        mostrarAlerta("Cliente registrado con éxito", "exito");
      }

      limpiarFormulario();
      setModalAbierto(false);
      await cargarTodo();
    } catch (error) {
      console.error("Error detallado en la operación:", error);
      mostrarAlerta("Error en base de datos: " + (error.message || error), "error");
    }
  };

  const prepararEdicion = (item) => {
    setCliForm({
      id_editando: item.id,
      nombre: item.nombre_completo || '',
      nit: item.nit_cc || '',
      tel: item.telefono || '', 
      email: item.correo || '',  
      dir: item.direccion || '',
      ciudad: item.ciudad || '',
      nota: item.nota || ''
    });
    setModalAbierto(true);
  };

  const eliminarCliente = async (id, nombre) => {
    if (window.confirm(`¿Estás seguro de inactivar al cliente "${nombre}"?`)) {
      try {
        const { error } = await supabase.from('clientes').update({ activo: false }).eq('id', id);
        if (error) throw error;
        
        mostrarAlerta("Cliente inactivado del directorio activo", "exito");
        await cargarTodo();
      } catch (err) {
        console.error("Error al inactivar:", err);
        mostrarAlerta("No se pudo inactivar el cliente", "error");
      }
    }
  };

  const limpiarFormulario = () => {
    setCliForm({ id_editando: null, nombre: '', nit: '', tel: '', email: '', dir: '', ciudad: '', nota: '' });
  };

  const abrirModalNuevo = () => {
    limpiarFormulario();
    setModalAbierto(true);
  };

  const exportarClientesAExcel = async () => {
    if (!lista || lista.length === 0) {
      mostrarAlerta("No hay clientes para exportar", "error");
      return;
    }

    try {
      const workbook = new ExcelJS.Workbook();
      const sheet = workbook.addWorksheet('Directorio Clientes');

      sheet.columns = [
        { header: 'NOMBRE COMPLETO', key: 'nombre', width: 30 },
        { header: 'NIT / CÉDULA', key: 'nit', width: 18 },
        { header: 'TELÉFONO / CELULAR', key: 'tel', width: 18 },
        { header: 'CIUDAD', key: 'ciudad', width: 20 },
        { header: 'CORREO ELECTRÓNICO', key: 'correo', width: 28 },
        { header: 'DIRECCIÓN', key: 'dir', width: 30 },
        { header: 'NOTAS INTERNAS', key: 'nota', width: 35 }
      ];

      lista.forEach(c => {
        sheet.addRow({
          nombre: (c.nombre_completo || '').toUpperCase(),
          nit: c.nit_cc || 'N/R',
          tel: c.telefono || 'N/R',
          ciudad: (c.ciudad || 'N/R').toUpperCase(),
          correo: c.correo || '---',
          dir: (c.direccion || 'N/R').toUpperCase(),
          nota: c.nota || '---'
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
          if ([2, 3, 4].includes(colN)) cell.alignment = { vertical: 'middle', horizontal: 'center' };
          else cell.alignment = { vertical: 'middle', horizontal: 'left' };
        });
      });

      sheet.autoFilter = { from: { row: 1, column: 1 }, to: { row: sheet.rowCount, column: sheet.columnCount } };

      const buffer = await workbook.xlsx.writeBuffer();
      const fechaHoy = new Date().toISOString().split('T')[0];
      const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      saveAs(blob, `DIRECTORIO_CLIENTES_${fechaHoy}.xlsx`);

      mostrarAlerta("Directorio de clientes exportado a Excel con éxito", "exito");
    } catch (err) {
      console.error("Error al exportar clientes:", err);
      mostrarAlerta("Error al generar el archivo Excel", "error");
    }
  };

  const clientesFiltrados = (lista || []).filter(c => {
    if (c.activo === false) return false;
    const query = busqueda.toLowerCase();
    return (
      (c.nombre_completo || '').toLowerCase().includes(query) ||
      (c.nit_cc || '').toLowerCase().includes(query) ||
      (c.ciudad || '').toLowerCase().includes(query) ||
      (c.telefono || '').toLowerCase().includes(query)
    );
  });

  return (
    <div className="space-y-6 pb-20 text-slate-800 dark:text-slate-200 font-sans transition-colors duration-300">
      
      {/* 🚀 CABECERA PRINCIPAL MODERNA */}
      <div className="bg-white dark:bg-slate-800 p-6 rounded-3xl shadow-xl border border-slate-200 dark:border-slate-700 flex flex-col md:flex-row justify-between items-center gap-4 transition-colors duration-300">
        <div className="space-y-1">
          <div className="flex items-center gap-3">
            <span className="text-2xl p-2 bg-green-700/10 dark:bg-emerald-500/20 rounded-xl text-green-700 dark:text-emerald-400">👥</span>
            <h2 className="text-xl font-black tracking-tight text-slate-900 dark:text-white uppercase">Directorio de Clientes</h2>
          </div>
          <p className="text-xs text-slate-500 dark:text-slate-400 font-medium">Gestión de base de datos comercial y contactos de ventas.</p>
        </div>
        
        <div className="flex items-center gap-3 w-full md:w-auto">
          <button onClick={exportarClientesAExcel} className="flex-1 md:flex-none px-4 py-2.5 bg-emerald-600 hover:bg-emerald-700 text-white font-black text-xs uppercase tracking-wider rounded-xl shadow-md transition-all flex items-center justify-center gap-2 cursor-pointer">
            <span>📊</span> Exportar Excel
          </button>
          <button onClick={abrirModalNuevo} className="flex-1 md:flex-none px-5 py-2.5 bg-green-700 hover:bg-green-800 text-white font-black text-xs uppercase tracking-wider rounded-xl shadow-md transition-all flex items-center justify-center gap-2 cursor-pointer">
            <span>+</span> Nuevo Cliente
          </button>
        </div>
      </div>

      {/* 🔍 BARRA DE BÚSQUEDA Y TOTALES */}
      <div className="bg-white dark:bg-slate-800 p-4 rounded-2xl shadow-md border border-slate-200 dark:border-slate-700 flex flex-col md:flex-row justify-between items-center gap-4 transition-colors duration-300">
        <div className="flex items-center gap-3">
          <span className="px-4 py-2 bg-slate-100 dark:bg-slate-700 text-slate-700 dark:text-slate-300 font-black text-[10px] uppercase rounded-xl border border-slate-200 dark:border-slate-600">
            Total Clientes: {clientesFiltrados.length}
          </span>
          <span className="px-3 py-1.5 bg-green-100 dark:bg-emerald-950/60 text-green-800 dark:text-emerald-400 font-black text-[10px] uppercase rounded-lg border border-green-200 dark:border-emerald-800">
            Comercial
          </span>
        </div>
        
        <div className="relative w-full md:w-96">
          <span className="absolute inset-y-0 left-0 flex items-center pl-3 text-slate-400">🔍</span>
          <input 
            type="text" 
            placeholder="Buscar por nombre, NIT, ciudad..." 
            className="w-full bg-slate-50 dark:bg-slate-700/50 border border-slate-200 dark:border-slate-600 text-xs text-slate-800 dark:text-white rounded-xl pl-10 pr-4 py-2.5 outline-none focus:border-green-700 font-bold placeholder-slate-400"
            value={busqueda}
            onChange={e => setBusqueda(e.target.value)}
          />
        </div>
      </div>

      {/* 📊 TABLA DE CLIENTES (ANCHO COMPLETO) */}
      <div className="bg-white dark:bg-slate-800 rounded-3xl shadow-xl overflow-hidden border border-slate-200 dark:border-slate-700 transition-colors duration-300">
        <div className="overflow-x-auto max-h-[600px]">
          <table className="w-full text-left text-xs border-collapse">
            <thead>
              <tr className="bg-slate-100 dark:bg-slate-700/80 text-slate-700 dark:text-slate-300 uppercase font-black text-[10px] tracking-wider border-b border-slate-200 dark:border-slate-700 sticky top-0">
                <th className="p-4">Nombre Cliente</th>
                <th className="p-4">NIT / Identificación</th>
                <th className="p-4">Contacto Directo</th>
                <th className="p-4 text-center">Ciudad</th>
                <th className="p-4 text-center">Acciones</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100 dark:divide-slate-700/60 font-bold text-slate-700 dark:text-slate-300">
              {clientesFiltrados.length === 0 ? (
                <tr>
                  <td colSpan="5" className="p-8 text-center text-slate-400 dark:text-slate-500 font-bold italic">
                    No hay clientes registrados o coincidentes con la búsqueda.
                  </td>
                </tr>
              ) : (
                clientesFiltrados.map((item, index) => {
                  const numLimpio = item.telefono ? item.telefono.replace(/\D/g, '') : '';
                  const linkWhatsApp = numLimpio ? `https://wa.me/57${numLimpio}` : null;

                  return (
                    <tr key={item.id} className={`${index % 2 === 0 ? 'bg-white dark:bg-slate-800' : 'bg-slate-50/50 dark:bg-slate-800/60'} hover:bg-sky-50/50 dark:hover:bg-slate-700/50 transition-colors border-l-8 border-green-700 dark:border-emerald-600`}>
                      
                      <td className="p-4 whitespace-nowrap">
                        <p className="uppercase text-sm font-black text-slate-900 dark:text-white">{item.nombre_completo}</p>
                        <p className="text-[10px] text-blue-600 dark:text-sky-400 font-bold lowercase mt-0.5">{item.correo || 'Sin correo electrónico'}</p>
                      </td>
                      
                      <td className="p-4 font-black text-slate-700 dark:text-slate-300 whitespace-nowrap">
                        {item.nit_cc}
                      </td>
                      
                      <td className="p-4 font-bold text-slate-700 dark:text-slate-300">
                        {linkWhatsApp ? (
                          <a href={linkWhatsApp} target="_blank" rel="noopener noreferrer" className="text-emerald-700 dark:text-emerald-400 hover:text-emerald-800 transition-colors flex items-center gap-1.5 font-black">
                            <span>💬</span> {item.telefono}
                          </a>
                        ) : (
                          <span className="text-slate-400 dark:text-slate-500 flex items-center gap-1.5"><span>📞</span> N/R</span>
                        )}
                        {item.direccion && (
                          <p className="text-[10px] text-slate-500 dark:text-slate-400 font-bold uppercase mt-1">🏠 {item.direccion}</p>
                        )}
                      </td>

                      <td className="p-4 text-center whitespace-nowrap">
                        <span className="inline-block bg-slate-100 dark:bg-slate-700 border border-slate-200 dark:border-slate-600 text-slate-700 dark:text-slate-300 px-2.5 py-1 rounded-md text-[9px] font-black uppercase shadow-sm">
                          📍 {item.ciudad || 'N/R'}
                        </span>
                      </td>

                      <td className="p-4 whitespace-nowrap">
                        <div className="flex gap-1.5 justify-center">
                          <button
                            type="button"
                            onClick={() => prepararEdicion(item)}
                            className="px-2.5 py-1.5 bg-slate-700 dark:bg-slate-600 text-white hover:bg-slate-900 dark:hover:bg-slate-500 rounded-lg shadow-sm border border-slate-800 dark:border-slate-500 transition-colors text-[10px] font-black flex items-center gap-1 cursor-pointer"
                            title="Editar Cliente"
                          >
                            <span>✏️</span> EDITAR
                          </button>
                          
                          <button
                            type="button"
                            onClick={() => eliminarCliente(item.id, item.nombre_completo)}
                            className="p-1.5 bg-red-100 dark:bg-red-950/60 text-red-700 dark:text-red-400 hover:bg-red-700 hover:text-white rounded-lg border border-red-200 dark:border-red-900 transition-colors cursor-pointer"
                            title="Inactivar Cliente"
                          >
                            🗑️
                          </button>
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

      {/* 🧊 MODAL FLOTANTE PARA NUEVO / EDITAR CLIENTE */}
      {modalAbierto && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 backdrop-blur-sm p-4 overflow-y-auto">
          <div className="bg-white dark:bg-slate-800 border border-slate-200 dark:border-slate-700 w-full max-w-2xl rounded-3xl shadow-2xl overflow-hidden animate-in fade-in zoom-in-95 duration-200 my-8 text-slate-800 dark:text-slate-200">
            
            <div className="p-5 bg-slate-900 text-white border-b border-slate-800 flex justify-between items-center">
              <div className="flex items-center gap-3">
                <span className="text-xl">👥</span>
                <div>
                  <h3 className="text-base font-black uppercase tracking-tight">
                    {cliForm.id_editando ? "Editar Cliente" : "Nuevo Cliente"}
                  </h3>
                  <p className="text-[10px] text-slate-400 font-medium">Información comercial y datos de contacto.</p>
                </div>
              </div>
              <button onClick={() => setModalAbierto(false)} className="text-slate-400 hover:text-white text-lg font-black px-3 py-1 rounded-xl bg-slate-800 cursor-pointer">✕</button>
            </div>

            <div className="p-6 space-y-4 max-h-[80vh] overflow-y-auto">
              
              <div>
                <label className="text-[10px] font-black text-slate-500 dark:text-slate-400 uppercase tracking-wider">Nombre Completo / Razón Social *</label>
                <input 
                  type="text" 
                  className="w-full bg-slate-50 dark:bg-slate-900 border border-slate-200 dark:border-slate-700 text-slate-800 dark:text-white p-2.5 rounded-xl font-bold text-sm uppercase outline-none focus:border-green-700 mt-1 placeholder-slate-400" 
                  value={cliForm.nombre || ''} 
                  onChange={e => setCliForm({...cliForm, nombre: e.target.value})} 
                  placeholder="Ej: JUAN PEREZ" 
                />
              </div>

              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                <div>
                  <label className="text-[10px] font-black text-slate-500 dark:text-slate-400 uppercase tracking-wider">NIT / Cédula *</label>
                  <input 
                    type="text" 
                    className="w-full bg-slate-50 dark:bg-slate-900 border border-slate-200 dark:border-slate-700 text-slate-800 dark:text-white p-2.5 rounded-xl font-bold text-sm outline-none focus:border-green-700 mt-1 placeholder-slate-400" 
                    value={cliForm.nit || ''} 
                    onChange={e => setCliForm({...cliForm, nit: e.target.value})} 
                    placeholder="900.000.000-1" 
                  />
                </div>
                <div>
                  <label className="text-[10px] font-black text-slate-500 dark:text-slate-400 uppercase tracking-wider">Correo Electrónico</label>
                  <input 
                    type="email" 
                    className="w-full bg-slate-50 dark:bg-slate-900 border border-slate-200 dark:border-slate-700 text-slate-800 dark:text-white p-2.5 rounded-xl font-bold text-xs lowercase outline-none focus:border-green-700 mt-1 placeholder-slate-400" 
                    value={cliForm.email || ''} 
                    onChange={e => setCliForm({...cliForm, email: e.target.value})} 
                    placeholder="cliente@correo.com" 
                  />
                </div>
              </div>

              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                <div>
                  <label className="text-[10px] font-black text-slate-500 dark:text-slate-400 uppercase tracking-wider">Teléfono / Celular</label>
                  <input 
                    type="text" 
                    className="w-full bg-slate-50 dark:bg-slate-900 border border-slate-200 dark:border-slate-700 text-slate-800 dark:text-white p-2.5 rounded-xl font-bold text-xs outline-none focus:border-green-700 mt-1 placeholder-slate-400" 
                    value={cliForm.tel || ''} 
                    onChange={e => setCliForm({...cliForm, tel: e.target.value})} 
                    placeholder="310..." 
                  />
                </div>
                <div>
                  <label className="text-[10px] font-black text-slate-500 dark:text-slate-400 uppercase tracking-wider">Ciudad</label>
                  <input 
                    type="text" 
                    className="w-full bg-slate-50 dark:bg-slate-900 border border-slate-200 dark:border-slate-700 text-slate-800 dark:text-white p-2.5 rounded-xl font-bold text-xs uppercase outline-none focus:border-green-700 mt-1 placeholder-slate-400" 
                    value={cliForm.ciudad || ''} 
                    onChange={e => setCliForm({...cliForm, ciudad: e.target.value})} 
                    placeholder="Bogotá" 
                  />
                </div>
              </div>

              <div>
                <label className="text-[10px] font-black text-slate-500 dark:text-slate-400 uppercase tracking-wider">Dirección Físíca</label>
                <input 
                  type="text" 
                  className="w-full bg-slate-50 dark:bg-slate-900 border border-slate-200 dark:border-slate-700 text-slate-800 dark:text-white p-2.5 rounded-xl font-bold text-xs uppercase outline-none focus:border-green-700 mt-1 placeholder-slate-400" 
                  value={cliForm.dir || ''} 
                  onChange={e => setCliForm({...cliForm, dir: e.target.value})} 
                  placeholder="Calle 10 #..." 
                />
              </div>

              <div>
                <label className="text-[10px] font-black text-slate-500 dark:text-slate-400 uppercase tracking-wider">Notas Internas / Preferencias</label>
                <textarea 
                  className="w-full bg-slate-50 dark:bg-slate-900 border border-slate-200 dark:border-slate-700 text-slate-800 dark:text-white p-2.5 rounded-xl font-bold h-20 text-xs outline-none focus:border-green-700 mt-1 resize-none placeholder-slate-400" 
                  value={cliForm.nota || ''} 
                  onChange={e => setCliForm({...cliForm, nota: e.target.value})} 
                  placeholder="Observaciones de entregas o contacto..." 
                />
              </div>

              <div className="flex justify-end gap-3 pt-4 border-t border-slate-200 dark:border-slate-700">
                <button type="button" onClick={() => setModalAbierto(false)} className="px-5 py-2.5 bg-slate-200 dark:bg-slate-700 hover:bg-slate-300 dark:hover:bg-slate-600 text-slate-700 dark:text-slate-200 font-black text-xs uppercase tracking-wider rounded-xl transition-colors cursor-pointer">
                  Cancelar
                </button>
                <button type="button" onClick={handleSave} className="px-6 py-2.5 bg-green-700 hover:bg-green-800 text-white font-black text-xs uppercase tracking-wider rounded-xl shadow-lg transition-colors cursor-pointer">
                  {cliForm.id_editando ? '💾 Guardar Cambios' : '🚀 Registrar Cliente'}
                </button>
              </div>

            </div>
          </div>
        </div>
      )}

    </div>
  );
}