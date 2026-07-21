import React from 'react';

export default function ConfigProv({ provForm, setProvForm, mostrarAlerta, cargarTodo, supabase, lista }) {
  
  // 🇨🇴 Se añade "Bre-B" a las opciones de pago en Colombia
  const bancosSoportados = ["Efectivo", "Bre-B (Pago Inmediato)", "Bancolombia Ahorros", "Bancolombia Corriente", "Nequi", "Daviplata", "Banco de Bogotá", "Colpatria", "Davivienda", "Otro"];

  const handleSave = async () => {
    if (!provForm.nombre || !provForm.nit) {
      mostrarAlerta("Nombre y NIT/CC son obligatorios", "error");
      return;
    }

    // Armamos el payload con el mapeo exacto incluyendo los campos bancarios
    const payload = {
      nombre: provForm.nombre.toUpperCase().trim(),
      nit_cc: provForm.nit.toString().trim(),
      telefono: provForm.tel ? provForm.tel.toString().trim() : null,
      direccion: provForm.dir ? provForm.dir.toUpperCase().trim() : null,
      ciudad: provForm.ciudad ? provForm.ciudad.toUpperCase().trim() : null,
      nota: provForm.nota ? provForm.nota.trim() : null,
      banco: provForm.banco || "Efectivo", 
      numero_cuenta: provForm.numero_cuenta ? provForm.numero_cuenta.trim() : null 
    };

    try {
      if (provForm.id_editando) {
        // === MODO EDICIÓN BLINDADO ===
        const { error: updateError } = await supabase
          .from('proveedores')
          .update(payload)
          .eq('id', provForm.id_editando);

        if (updateError) throw updateError;
        mostrarAlerta("Proveedor actualizado con éxito", "exito");
      } else {
        // === MODO CREACIÓN ===
        const { error: insertError } = await supabase
          .from('proveedores')
          .insert([payload]);

        if (insertError) throw insertError;
        mostrarAlerta("Proveedor guardado correctamente", "exito");
      }

      limpiarFormulario();
      await cargarTodo();
    } catch (error) {
      console.error("Error en la operación de proveedores:", error);
      mostrarAlerta("Error en base de datos: " + (error.message || error), "error");
    }
  };

  const prepararEdicion = (item) => {
    setProvForm({
      id_editando: item.id,
      nombre: item.nombre || '',
      nit: item.nit_cc || '',
      tel: item.telefono || '',
      dir: item.direccion || '',
      ciudad: item.ciudad || '',
      nota: item.nota || '',
      banco: item.banco || 'Efectivo', 
      numero_cuenta: item.numero_cuenta || '' 
    });
    window.scrollTo({ top: 0, behavior: 'smooth' });
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
      dir: '', 
      ciudad: '', 
      nota: '',
      banco: 'Efectivo',
      numero_cuenta: '' 
    });
  };

  return (
    <div className="space-y-6 pb-20">
      <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
        
        {/* COLUMNA IZQUIERDA: FORMULARIO (1/3) - Unificado a Azul Océano */}
        <div className="bg-white p-6 rounded-3xl shadow-xl border-t-8 border-[#117097] h-fit">
          <h3 className="font-black text-slate-800 uppercase text-xs italic mb-5">
            {provForm.id_editando ? '📝 Editar Proveedor' : '🚛 Nuevo Proveedor'}
          </h3>
          
          <div className="space-y-4">
            <div>
              <label className="text-[10px] font-bold text-gray-400 uppercase px-1 italic">Razon Social / Nombre</label>
              <input className="w-full border-2 p-2.5 rounded-xl font-bold text-sm uppercase outline-none focus:border-[#117097]" value={provForm.nombre || ''} onChange={e=>setProvForm({...provForm, nombre: e.target.value})} placeholder="Ej: AGROINSUMOS SAS" />
            </div>

            <div>
              <label className="text-[10px] font-bold text-gray-400 uppercase px-1 italic">NIT / Cédula</label>
              <input className="w-full border-2 p-2.5 rounded-xl font-bold text-sm outline-none focus:border-[#117097]" value={provForm.nit || ''} onChange={e=>setProvForm({...provForm, nit: e.target.value})} placeholder="800.000.000-1" />
            </div>

            <div className="grid grid-cols-2 gap-2">
              <div>
                <label className="text-[10px] font-bold text-gray-400 uppercase px-1 italic">Teléfono / Celular</label>
                <input className="w-full border-2 p-2.5 rounded-xl font-bold text-xs outline-none focus:border-[#117097]" value={provForm.tel || ''} onChange={e=>setProvForm({...provForm, tel: e.target.value})} placeholder="315..." />
              </div>
              <div>
                <label className="text-[10px] font-bold text-gray-400 uppercase px-1 italic">Ciudad</label>
                <input className="w-full border-2 p-2.5 rounded-xl font-bold text-xs uppercase outline-none focus:border-[#117097]" value={provForm.ciudad || ''} onChange={e=>setProvForm({...provForm, ciudad: e.target.value})} placeholder="Chía" />
              </div>
            </div>

            <div>
              <label className="text-[10px] font-bold text-gray-400 uppercase px-1 italic">Dirección Física</label>
              <input className="w-full border-2 p-2.5 rounded-xl font-bold text-xs uppercase outline-none focus:border-[#117097]" value={provForm.dir || ''} onChange={e=>setProvForm({...provForm, dir: e.target.value})} placeholder="Zona Industrial..." />
            </div>

            {/* 💳 SECCIÓN: MEDIOS DE PAGO Y CUENTA */}
            <div className="bg-slate-50 p-3 rounded-2xl border-2 border-slate-100 space-y-3 shadow-inner">
              <p className="text-[9px] font-black text-slate-500 uppercase tracking-wider border-b pb-1">Información de Pago / Transferencia</p>
              
              <div className="grid grid-cols-1 md:grid-cols-2 gap-2">
                <div>
                  <label className="text-[9px] font-black text-slate-500 uppercase px-1">Forma / Banco</label>
                  <select 
                    className="w-full p-2 border-2 bg-white rounded-xl outline-none text-xs font-bold focus:border-[#117097]" 
                    value={provForm.banco || 'Efectivo'} 
                    onChange={e => setProvForm({...provForm, banco: e.target.value})}
                  >
                    {bancosSoportados.map(b => <option key={b} value={b}>{b}</option>)}
                  </select>
                </div>
                <div>
                  <label className="text-[9px] font-black text-slate-500 uppercase px-1">N° Cuenta / Llave Bre-B</label>
                  <input 
                    type="text" 
                    className="w-full p-2 border-2 bg-white rounded-xl outline-none text-xs font-black text-slate-700 focus:border-[#117097]" 
                    value={provForm.numero_cuenta || ''} 
                    onChange={e => setProvForm({...provForm, numero_cuenta: e.target.value})} 
                    placeholder="Celular, cédula o cuenta..." 
                  />
                </div>
              </div>
            </div>

            <div>
              <label className="text-[10px] font-bold text-gray-400 uppercase px-1 italic">Observaciones / Notas</label>
              <textarea className="w-full border-2 p-2.5 rounded-xl font-bold h-20 text-xs outline-none focus:border-[#117097]" value={provForm.nota || ''} onChange={e=>setProvForm({...provForm, nota: e.target.value})} placeholder="Detalles de entrega, horarios..." />
            </div>

            <div className="flex gap-2 pt-2">
              <button onClick={handleSave} className={`flex-1 py-3.5 text-white font-black rounded-xl uppercase text-xs tracking-widest shadow-md transition-all ${provForm.id_editando ? 'bg-amber-600 hover:bg-amber-700' : 'bg-[#117097] hover:bg-[#0a4c68]'}`}>
                {provForm.id_editando ? '💾 Actualizar' : '💾 Guardar'}
              </button>
              {provForm.id_editando && (
                <button onClick={limpiarFormulario} className="px-4 py-3.5 bg-gray-200 text-gray-700 font-black rounded-xl uppercase text-xs hover:bg-gray-300">X</button>
              )}
            </div>
          </div>
        </div>

        {/* COLUMNA DERECHA: TABLA DE REGISTROS (2/3) - Unificada a Azul Océano */}
        <div className="lg:col-span-2 bg-white rounded-3xl shadow-xl overflow-hidden border border-gray-200">
          <div className="p-4 bg-slate-800 text-white font-black text-xs uppercase tracking-widest italic flex justify-between items-center">
            <span>Base de Datos Proveedores</span>
            <span className="text-[10px] bg-[#117097] px-2 py-0.5 rounded-md font-bold uppercase">Logística</span>
          </div>
          <div className="overflow-x-auto">
            <table className="w-full text-left text-[11px] border-collapse">
              <thead>
                <tr className="bg-gray-300 text-slate-800 uppercase font-black">
                  <th className="p-4 border-b-2 border-gray-400">Nombre / Razón Social</th>
                  <th className="p-4 border-b-2 border-gray-400">NIT / CC</th>
                  <th className="p-4 border-b-2 border-gray-400">Datos de Pago</th>
                  <th className="p-4 border-b-2 border-gray-400">Contacto / Ubicación</th>
                  <th className="p-4 border-b-2 border-gray-400 text-center">Acciones</th>
                </tr>
              </thead>
              <tbody className="divide-y-2 divide-gray-400">
                {lista?.length === 0 ? (
                  <tr>
                    <td colSpan="5" className="p-6 text-center text-gray-400 italic font-bold">No hay proveedores registrados.</td>
                  </tr>
                ) : (
                  lista?.map((item, index) => (
                    <tr key={item.id} className={`${index % 2 === 0 ? 'bg-white' : 'bg-gray-100'} hover:bg-sky-50 transition-colors`}>
                      <td className="p-4 font-black text-slate-900 border-l-8 border-[#117097]">
                        <p className="uppercase">{item.nombre}</p>
                        {item.nota && <p className="text-[9px] text-gray-400 font-medium normal-case mt-0.5 line-clamp-1">📝 {item.nota}</p>}
                      </td>
                      <td className="p-4 font-bold text-slate-700">{item.nit_cc}</td>
                      
                      {/* 💳 COLUMNA DE DATOS DE PAGO CON NUEVO DISTINTIVO TURQUESA PARA BRE-B */}
                      <td className="p-4 font-bold">
                        <span className={`inline-block px-2 py-0.5 rounded text-[8px] uppercase tracking-wider font-black ${
                          item.banco === 'Efectivo' 
                            ? 'bg-emerald-100 text-emerald-800' 
                            : item.banco?.toLowerCase().includes('bre-b')
                            ? 'bg-cyan-100 text-cyan-800 border border-cyan-300' // 👈 Estilo para Bre-B
                            : item.banco?.toLowerCase().includes('nequi') 
                            ? 'bg-purple-100 text-purple-800' 
                            : item.banco?.toLowerCase().includes('daviplata')
                            ? 'bg-red-100 text-red-800'
                            : 'bg-blue-100 text-blue-800'
                        }`}>
                          {item.banco?.toLowerCase().includes('bre-b') ? '⚡ Bre-B' : `🏦 ${item.banco || 'Efectivo'}`}
                        </span>
                        {item.numero_cuenta && (
                          <p className="text-[10px] text-slate-600 mt-1 font-black">
                            {item.banco?.toLowerCase().includes('bre-b') ? `Llave: ${item.numero_cuenta}` : `#${item.numero_cuenta}`}
                          </p>
                        )}
                      </td>

                      <td className="p-4 font-bold text-slate-600">
                        <p>📞 {item.telefono || 'N/R'}</p>
                        <p className="uppercase text-slate-400 text-[9px] mt-0.5">📍 {item.ciudad || 'N/R'}</p>
                      </td>
                      <td className="p-4">
                        <div className="flex justify-center gap-2">
                          <button onClick={() => prepararEdicion(item)} className="p-2 bg-amber-100 text-amber-700 rounded-lg hover:bg-amber-700 hover:text-white transition-colors border border-amber-200">
                            ✏️
                          </button>
                          <button onClick={() => eliminarProveedor(item.id, item.nombre)} className="p-2 bg-red-100 text-red-700 rounded-lg hover:bg-red-700 hover:text-white transition-colors border border-red-200">
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

      </div>
    </div>
  );
}