import React, { useState } from 'react';
import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';

export default function ConfigCli({ cliForm, setCliForm, mostrarAlerta, cargarTodo, supabase, lista }) {
  const [busqueda, setBusqueda] = useState('');

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
    window.scrollTo({ top: 0, behavior: 'smooth' });
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

  // --- 📊 EXPORTAR DIRECTORIO COMERCIAL DE CLIENTES A EXCEL ---
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

      // Estilo de Encabezado
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

  // Filtrado de búsqueda
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
    <div className="space-y-6 pb-20 text-slate-800">
      <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
        
        {/* COLUMNA IZQUIERDA: FORMULARIO */}
        <div className="bg-white p-6 rounded-3xl shadow-xl border-t-8 border-green-800 h-fit">
          <h3 className="font-black text-slate-800 uppercase text-xs italic mb-5">
            {cliForm.id_editando ? '✏️ Editar Cliente' : '👤 Nuevo Cliente'}
          </h3>
          
          <div className="space-y-4">
            <div>
              <label className="text-[10px] font-bold text-gray-400 uppercase px-1 italic">Nombre Completo / Razón Social *</label>
              <input className="w-full border-2 p-2.5 rounded-xl font-bold text-sm uppercase outline-none focus:border-green-700" value={cliForm.nombre} onChange={e=>setCliForm({...cliForm, nombre: e.target.value})} placeholder="Ej: JUAN PEREZ" />
            </div>

            <div>
              <label className="text-[10px] font-bold text-gray-400 uppercase px-1 italic">NIT / Cédula *</label>
              <input className="w-full border-2 p-2.5 rounded-xl font-bold text-sm outline-none focus:border-green-700" value={cliForm.nit} onChange={e=>setCliForm({...cliForm, nit: e.target.value})} placeholder="900.000.000-1" />
            </div>

            <div className="grid grid-cols-2 gap-2">
              <div>
                <label className="text-[10px] font-bold text-gray-400 uppercase px-1 italic">Teléfono / Celular</label>
                <input className="w-full border-2 p-2.5 rounded-xl font-bold text-xs outline-none focus:border-green-700" value={cliForm.tel} onChange={e=>setCliForm({...cliForm, tel: e.target.value})} placeholder="310..." />
              </div>
              <div>
                <label className="text-[10px] font-bold text-gray-400 uppercase px-1 italic">Ciudad</label>
                <input className="w-full border-2 p-2.5 rounded-xl font-bold text-xs uppercase outline-none focus:border-green-700" value={cliForm.ciudad} onChange={e=>setCliForm({...cliForm, ciudad: e.target.value})} placeholder="Bogotá" />
              </div>
            </div>

            <div>
              <label className="text-[10px] font-bold text-gray-400 uppercase px-1 italic">Correo Electrónico</label>
              <input type="email" className="w-full border-2 p-2.5 rounded-xl font-bold text-xs outline-none focus:border-green-700" value={cliForm.email} onChange={e=>setCliForm({...cliForm, email: e.target.value})} placeholder="cliente@correo.com" />
            </div>

            <div>
              <label className="text-[10px] font-bold text-gray-400 uppercase px-1 italic">Dirección</label>
              <input className="w-full border-2 p-2.5 rounded-xl font-bold text-xs uppercase outline-none focus:border-green-700" value={cliForm.dir} onChange={e=>setCliForm({...cliForm, dir: e.target.value})} placeholder="Calle 10 #..." />
            </div>

            <div>
              <label className="text-[10px] font-bold text-gray-400 uppercase px-1 italic">Notas Internas / Preferencias</label>
              <textarea className="w-full border-2 p-2.5 rounded-xl font-bold h-16 text-xs outline-none focus:border-green-700" value={cliForm.nota} onChange={e=>setCliForm({...cliForm, nota: e.target.value})} placeholder="Observaciones..." />
            </div>

            <div className="flex gap-2 pt-2">
              <button onClick={handleSave} className={`flex-1 py-3.5 text-white font-black rounded-xl uppercase text-xs tracking-widest shadow-md transition-colors ${cliForm.id_editando ? 'bg-amber-600 hover:bg-amber-700' : 'bg-green-800 hover:bg-green-900'}`}>
                {cliForm.id_editando ? '💾 Actualizar' : '💾 Guardar Cliente'}
              </button>
              {cliForm.id_editando && (
                <button onClick={limpiarFormulario} className="px-4 py-3.5 bg-gray-200 text-gray-700 font-black rounded-xl uppercase text-xs hover:bg-gray-300">X</button>
              )}
            </div>
          </div>
        </div>

        {/* COLUMNA DERECHA: TABLA BUSCABLE + EXPORTAR EXCEL */}
        <div className="lg:col-span-2 bg-white rounded-3xl shadow-xl overflow-hidden border border-gray-200 flex flex-col">
          
          <div className="p-4 bg-slate-800 text-white font-black text-xs uppercase tracking-wider flex flex-col sm:flex-row justify-between items-center gap-3">
            <div className="flex items-center gap-2">
              <span>Base de Datos Clientes ({clientesFiltrados.length})</span>
              <span className="text-[9px] bg-green-700 px-2 py-0.5 rounded-md font-bold">Comercial</span>
            </div>

            <div className="flex items-center gap-2 flex-wrap w-full sm:w-auto justify-end">
              <button
                onClick={exportarClientesAExcel}
                className="px-3 py-1.5 bg-emerald-600 hover:bg-emerald-700 text-white text-[10px] font-black uppercase rounded-xl shadow transition-colors"
              >
                📊 EXPORTAR EXCEL
              </button>

              <input 
                type="text" 
                placeholder="🔍 Buscar cliente, NIT, ciudad..." 
                className="px-3 py-1.5 text-xs rounded-xl text-slate-800 outline-none font-bold placeholder-gray-400 min-w-[200px]"
                value={busqueda}
                onChange={e => setBusqueda(e.target.value)}
              />
            </div>
          </div>

          <div className="overflow-x-auto min-h-[400px]">
            <table className="w-full text-left text-[11px] border-collapse">
              <thead>
                <tr className="bg-gray-200 text-slate-800 uppercase font-black sticky top-0">
                  <th className="p-3.5 border-b border-gray-300">Nombre Cliente</th>
                  <th className="p-3.5 border-b border-gray-300">NIT / Identificación</th>
                  <th className="p-3.5 border-b border-gray-300">Contacto Directo</th>
                  <th className="p-3.5 border-b border-gray-300 text-center">Ciudad</th>
                  <th className="p-3.5 border-b border-gray-300 text-center">Acciones</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-gray-200 font-bold text-slate-700">
                {clientesFiltrados.length === 0 ? (
                  <tr>
                    <td colSpan="5" className="p-8 text-center text-gray-400 italic font-bold">No hay clientes registrados o coincidentes con la búsqueda.</td>
                  </tr>
                ) : (
                  clientesFiltrados.map((item, index) => {
                    const numLimpio = item.telefono ? item.telefono.replace(/\D/g, '') : '';
                    const linkWhatsApp = numLimpio ? `https://wa.me/57${numLimpio}` : null;

                    return (
                      <tr key={item.id} className={`${index % 2 === 0 ? 'bg-white' : 'bg-slate-50'} hover:bg-sky-50 transition-colors`}>
                        <td className="p-3.5 font-black text-slate-900 border-l-4 border-green-700">
                          <p className="uppercase text-xs">{item.nombre_completo}</p>
                          <p className="text-[9px] text-blue-600 lowercase font-bold">{item.correo || '---'}</p>
                        </td>

                        <td className="p-3.5 font-black text-slate-600">{item.nit_cc}</td>

                        <td className="p-3.5">
                          {linkWhatsApp ? (
                            <a href={linkWhatsApp} target="_blank" rel="noopener noreferrer" className="text-emerald-700 hover:underline flex items-center gap-1 font-black">
                              💬 {item.telefono}
                            </a>
                          ) : (
                            <span className="text-gray-400">📞 N/R</span>
                          )}
                          {item.direccion && <p className="text-[9px] text-gray-400 font-bold uppercase mt-0.5">🏠 {item.direccion}</p>}
                        </td>

                        <td className="p-3.5 text-center">
                          <span className="inline-block bg-slate-100 text-slate-700 px-2 py-0.5 rounded text-[9px] font-black uppercase">
                            📍 {item.ciudad || 'N/R'}
                          </span>
                        </td>

                        <td className="p-3.5 text-center">
                          <div className="flex justify-center gap-1.5">
                            <button onClick={() => prepararEdicion(item)} className="p-1.5 bg-amber-100 text-amber-700 rounded-lg hover:bg-amber-600 hover:text-white transition-all border border-amber-200 text-xs" title="Editar">
                              ✏️
                            </button>
                            <button onClick={() => eliminarCliente(item.id, item.nombre_completo)} className="p-1.5 bg-red-100 text-red-700 rounded-lg hover:bg-red-600 hover:text-white transition-all border border-red-200 text-xs" title="Inactivar">
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

      </div>
    </div>
  );
}