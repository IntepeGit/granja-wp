import React, { useState, useEffect } from 'react';
import { supabase } from '../lib/supabase';
import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';

export default function Inventario({ mostrarAlerta, datosInvernaderos, userRole }) {
  // Estados principales
  const [listaInventario, setListaInventario] = useState([]);
  const [listaTrabajadores, setListaTrabajadores] = useState([]);
  const [listaUbicaciones, setListaUbicaciones] = useState([]);
  const [nuevaUbicacionTexto, setNuevaUbicacionTexto] = useState('');

  const [tabActiva, setTabActiva] = useState('bodega'); // 'bodega', 'entrada', 'salida', 'reubicar'
  const [filtroTipo, setFiltroTipo] = useState('TODOS'); 
  const [cargando, setCargando] = useState(false);
  const [busqueda, setBusqueda] = useState('');

  // Edición directa
  const [idEditando, setIdEditando] = useState(null);
  const [filaEditable, setFilaEditable] = useState({ 
    nombre_insumo: '', categoria: '', unidad_medida: '', cantidad_actual: 0, 
    stock_minimo: 0, aplica_stock: true, tipo_item: 'Consumible', ubicacion: 'BODEGA PRINCIPAL', estado_herramienta: 'Operativo'
  });

  // Formularios
  const [insumoForm, setInsumoForm] = useState({ 
    nombre_insumo: '', tipo_item: 'Consumible', categoria: 'Fertilizante', 
    unidad_medida: 'Bulto', stock_minimo: 5, aplica_stock: true, ubicacion: 'BODEGA PRINCIPAL', estado_herramienta: 'Operativo'
  });
  
  const [entradaForm, setEntradaForm] = useState({ 
    insumo_id: '', cantidad_ingresada: '', precio_unitario_compra: '', numero_factura_comprobante: '', observaciones: '' 
  });
  
  const [salidaForm, setSalidaForm] = useState({ 
    insumo_id: '', cantidad_retirada: '', invernadero_id: '', responsable: '', nota_uso: '' 
  });

  const [reubicacionForm, setReubicacionForm] = useState({
    herramienta_id: '', nueva_ubicacion: '', responsable: '', observaciones: ''
  });

  const categorias = ['Fertilizante', 'Veneno', 'Semilla', 'Herramienta', 'Maquinaria', 'Canastilla', 'Plastico', 'Accesorio', 'Manguera', 'Otros'];
  const unidades = ['Bulto', 'Kilo', 'Litro', 'Metro', 'Unidad', 'Caja', 'Galón'];
  const estadosHerramientaOptions = ['Operativo', 'En Mantenimiento', 'En Préstamo', 'Dañado / Inactivo'];

  const esAdmin = userRole === 'admin';

  useEffect(() => {
    cargarInventario();
    cargarTrabajadoresNomina();
    cargarUbicaciones();
  }, []);

  const cargarInventario = async () => {
    setCargando(true);
    try {
      const { data, error } = await supabase.from('inventario').select('*').order('nombre_insumo', { ascending: true });
      if (error) throw error;
      setListaInventario(data || []);
    } catch (err) {
      if (mostrarAlerta) mostrarAlerta("No se pudo cargar el inventario", "error");
    } finally {
      setCargando(false);
    }
  };

  const cargarTrabajadoresNomina = async () => {
    try {
      const { data } = await supabase.from('nomina_trabajadores').select('id, nombre_completo').eq('activo', true).order('nombre_completo', { ascending: true });
      setListaTrabajadores(data || []);
    } catch (err) { console.error(err); }
  };

  const cargarUbicaciones = async () => {
    try {
      const { data } = await supabase.from('ubicaciones_inventario').select('*').order('nombre', { ascending: true });
      setListaUbicaciones(data || []);
    } catch (err) { console.error(err); }
  };

  const agregarNuevaUbicacion = async (e) => {
    e.preventDefault();
    if (!nuevaUbicacionTexto.trim()) return;
    try {
      const { error } = await supabase.from('ubicaciones_inventario').insert([{ nombre: nuevaUbicacionTexto.toUpperCase().trim() }]);
      if (error) throw error;
      mostrarAlerta("Ubicación física creada", "exito");
      setNuevaUbicacionTexto('');
      cargarUbicaciones();
    } catch (err) { mostrarAlerta("La ubicación ya existe o hubo un error", "error"); }
  };

  const crearInsumoBauche = async (e) => {
    e.preventDefault();
    if (!esAdmin || !insumoForm.nombre_insumo) return;

    try {
      const esConsumible = insumoForm.tipo_item === 'Consumible';
      const { error } = await supabase.from('inventario').insert([{
        nombre_insumo: insumoForm.nombre_insumo.toUpperCase().trim(),
        tipo_item: insumoForm.tipo_item,
        categoria: insumoForm.categoria,
        unidad_medida: insumoForm.unidad_medida,
        stock_minimo: esConsumible && insumoForm.aplica_stock ? (parseFloat(insumoForm.stock_minimo) || 0) : 0,
        cantidad_actual: esConsumible ? 0 : 1,
        aplica_stock: esConsumible ? insumoForm.aplica_stock : false,
        ubicacion: insumoForm.ubicacion.toUpperCase().trim(),
        estado_herramienta: insumoForm.estado_herramienta
      }]);

      if (error) throw error;
      mostrarAlerta("Artículo registrado en bodega", "exito");
      setInsumoForm({ nombre_insumo: '', tipo_item: 'Consumible', categoria: 'Fertilizante', unidad_medida: 'Bulto', stock_minimo: 5, aplica_stock: true, ubicacion: 'BODEGA PRINCIPAL', estado_herramienta: 'Operativo' });
      cargarInventario();
    } catch (err) { mostrarAlerta("El registro ya existe o hubo un error", "error"); }
  };

  const registrarEntradaBodega = async (e) => {
    e.preventDefault();
    const cant = parseFloat(entradaForm.cantidad_ingresada);
    const precio = parseFloat(entradaForm.precio_unitario_compra) || 0;
    if (!entradaForm.insumo_id || cant <= 0) return;

    try {
      const { error } = await supabase.from('entradas_inventario').insert([{
        insumo_id: entradaForm.insumo_id,
        cantidad_ingresada: cant,
        precio_unitario_compra: precio,
        monto_total_compra: cant * precio,
        numero_factura_comprobante: entradaForm.numero_factura_comprobante || 'S/N',
        observaciones: entradaForm.observaciones
      }]);
      if (error) throw error;
      mostrarAlerta("Ingreso registrado e inventario incrementado", "exito");
      setEntradaForm({ insumo_id: '', cantidad_ingresada: '', precio_unitario_compra: '', numero_factura_comprobante: '', observaciones: '' });
      cargarInventario();
    } catch (err) { mostrarAlerta("Error al procesar la entrada", "error"); }
  };

  const registrarSalidaInvernadero = async (e) => {
    e.preventDefault();
    const cant = parseFloat(salidaForm.cantidad_retirada);
    if (!salidaForm.insumo_id || cant <= 0 || !salidaForm.responsable) return;

    const insumoSeleccionado = listaInventario.find(i => i.id === salidaForm.insumo_id);
    if (insumoSeleccionado && insumoSeleccionado.cantidad_actual < cant) {
      mostrarAlerta("No puedes retirar más existencias de las disponibles", "error");
      return;
    }

    try {
      const { error } = await supabase.from('salidas_inventario').insert([{
        insumo_id: salidaForm.insumo_id,
        cantidad_retirada: cant,
        invernadero_id: salidaForm.invernadero_id || null,
        responsable: salidaForm.responsable.toUpperCase(),
        nota_uso: salidaForm.nota_uso
      }]);
      if (error) throw error;
      mostrarAlerta("Insumo consumido y descontado de bodega", "exito");
      setSalidaForm({ insumo_id: '', cantidad_retirada: '', invernadero_id: '', responsable: '', nota_uso: '' });
      cargarInventario();
    } catch (err) { mostrarAlerta("Error al registrar el consumo", "error"); }
  };

  const reubicarHerramienta = async (e) => {
    e.preventDefault();
    if (!reubicacionForm.herramienta_id || !reubicacionForm.nueva_ubicacion) return;

    try {
      const { error } = await supabase.from('inventario').update({
        ubicacion: reubicacionForm.nueva_ubicacion.toUpperCase().trim(),
        estado_herramienta: reubicacionForm.responsable ? `En uso por ${reubicacionForm.responsable.toUpperCase()}` : 'Operativo'
      }).eq('id', reubicacionForm.herramienta_id);

      if (error) throw error;
      mostrarAlerta("Ubicación de herramienta actualizada", "exito");
      setReubicacionForm({ herramienta_id: '', nueva_ubicacion: '', responsable: '', observaciones: '' });
      cargarInventario();
    } catch (err) { mostrarAlerta("Error al reubicar la herramienta", "error"); }
  };

  const iniciarEdicion = (ins) => {
    if (!esAdmin) return;
    setIdEditando(ins.id);
    setFilaEditable({
      nombre_insumo: ins.nombre_insumo,
      categoria: ins.categoria,
      unidad_medida: ins.unidad_medida,
      cantidad_actual: ins.cantidad_actual,
      stock_minimo: ins.stock_minimo,
      aplica_stock: ins.aplica_stock !== false,
      tipo_item: ins.tipo_item || 'Consumible',
      ubicacion: ins.ubicacion || 'BODEGA PRINCIPAL',
      estado_herramienta: ins.estado_herramienta || 'Operativo'
    });
  };

  const guardarEdicionFila = async (id) => {
    if (!esAdmin) return;
    try {
      const { error } = await supabase.from('inventario').update({
        nombre_insumo: filaEditable.nombre_insumo.toUpperCase().trim(),
        tipo_item: filaEditable.tipo_item,
        categoria: filaEditable.categoria,
        unidad_medida: filaEditable.unidad_medida,
        cantidad_actual: parseFloat(filaEditable.cantidad_actual) || 0,
        stock_minimo: parseFloat(filaEditable.stock_minimo) || 0,
        aplica_stock: filaEditable.aplica_stock,
        ubicacion: filaEditable.ubicacion.toUpperCase().trim(),
        estado_herramienta: filaEditable.estado_herramienta
      }).eq('id', id);

      if (error) throw error;
      mostrarAlerta("Registro actualizado con éxito", "exito");
      setIdEditando(null);
      cargarInventario();
    } catch (err) { mostrarAlerta("Error al actualizar", "error"); }
  };

  const eliminarInsumoCompleto = async (id, nombre) => {
    if (!esAdmin) return;
    if (window.confirm(`¿Está seguro de eliminar COMPLETAMENTE "${nombre}"?`)) {
      try {
        const { error } = await supabase.from('inventario').delete().eq('id', id);
        if (error) throw error;
        mostrarAlerta("Eliminado del sistema", "exito");
        cargarInventario();
      } catch (err) { mostrarAlerta("No se puede eliminar porque tiene historial", "error"); }
    }
  };

  const obtenerEstadoStock = (ins) => {
    if (ins.tipo_item !== 'Consumible' || ins.aplica_stock === false) {
      return { nivel: 'herramienta', etiqueta: ins.estado_herramienta || 'Activo', claseTabla: 'bg-slate-100 text-slate-800 border border-slate-300', claseTexto: 'text-slate-800' };
    }
    if (ins.cantidad_actual <= 2) {
      return { nivel: 'rojo', etiqueta: 'Crítico (≤2)', claseTabla: 'bg-red-500 text-white animate-pulse', claseTexto: 'text-red-600 font-black' };
    } else if (ins.cantidad_actual >= 3 && ins.cantidad_actual <= 5) {
      return { nivel: 'amarillo', etiqueta: 'Bajo (3-5)', claseTabla: 'bg-amber-400 text-slate-900', claseTexto: 'text-amber-600 font-black' };
    }
    return { nivel: 'verde', etiqueta: 'Disponible', claseTabla: 'bg-emerald-100 text-emerald-800', claseTexto: 'text-emerald-700' };
  };

  // --- 📊 EXPORTAR INVENTARIO A EXCEL EN 2 HOJAS (CONSUMIBLES Y HERRAMIENTAS) ---
  const exportarInventarioAExcel = async () => {
    if (!listaInventario || listaInventario.length === 0) {
      mostrarAlerta("No hay artículos en bodega para exportar", "error");
      return;
    }

    try {
      const workbook = new ExcelJS.Workbook();

      const itemsConsumibles = listaInventario.filter(i => i.tipo_item === 'Consumible');
      const itemsHerramientas = listaInventario.filter(i => i.tipo_item !== 'Consumible');

      // 1. HOJA DE CONSUMIBLES
      const sheetConsumibles = workbook.addWorksheet('Consumibles');
      sheetConsumibles.columns = [
        { header: 'ARTÍCULO / INSUMO', key: 'nombre', width: 30 },
        { header: 'CATEGORÍA', key: 'categoria', width: 20 },
        { header: 'UBICACIÓN FÍSICA', key: 'ubicacion', width: 25 },
        { header: 'CANTIDAD ACTUAL', key: 'cantidad', width: 18 },
        { header: 'UNIDAD MEDIDA', key: 'unidad', width: 16 },
        { header: 'CONTROL STOCK', key: 'aplica_stock', width: 16 },
        { header: 'STOCK MÍNIMO', key: 'stock_minimo', width: 16 },
        { header: 'ESTADO DISPONIBILIDAD', key: 'estado', width: 22 }
      ];

      itemsConsumibles.forEach(c => {
        const est = obtenerEstadoStock(c);
        sheetConsumibles.addRow({
          nombre: c.nombre_insumo.toUpperCase(),
          categoria: (c.categoria || 'GENERAL').toUpperCase(),
          ubicacion: (c.ubicacion || 'BODEGA PRINCIPAL').toUpperCase(),
          cantidad: parseFloat(c.cantidad_actual) || 0,
          unidad: c.unidad_medida,
          aplica_stock: c.aplica_stock !== false ? 'SÍ' : 'NO',
          stock_minimo: c.stock_minimo || 0,
          estado: est.etiqueta.toUpperCase()
        });
      });

      const ultFilaC = sheetConsumibles.rowCount;
      const totalRowC = sheetConsumibles.addRow({
        ubicacion: 'TOTAL EXISTENCIAS:',
        cantidad: { formula: `=SUM(D2:D${ultFilaC})` }
      });

      // Estilizar Hoja 1
      const headC = sheetConsumibles.getRow(1);
      headC.height = 24;
      headC.eachCell(cell => {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF117097' } };
        cell.font = { name: 'Arial', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
      });

      sheetConsumibles.eachRow((row, rNum) => {
        if (rNum === 1 || rNum === ultFilaC + 1) return;
        row.height = 20;
        const cebra = rNum % 2 === 0;
        row.eachCell((cell, colN) => {
          cell.font = { name: 'Arial', size: 9 };
          if (cebra) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEBF5FB' } };
          if ([2, 3, 5, 6, 8].includes(colN)) cell.alignment = { vertical: 'middle', horizontal: 'center' };
          else if ([4, 7].includes(colN)) cell.alignment = { vertical: 'middle', horizontal: 'right' };
          else cell.alignment = { vertical: 'middle', horizontal: 'left' };
        });
      });

      totalRowC.height = 22;
      totalRowC.eachCell((cell, colN) => {
        cell.font = { name: 'Arial', size: 10, bold: true, color: { argb: 'FF0A4C68' } };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD6EEFC' } };
        cell.border = { top: { style: 'thin' }, bottom: { style: 'double' } };
        if (colN === 3) cell.alignment = { vertical: 'middle', horizontal: 'right' };
        if (colN === 4) cell.alignment = { vertical: 'middle', horizontal: 'right' };
      });

      // 2. HOJA DE HERRAMIENTAS Y ACTIVOS FIJOS
      const sheetHerramientas = workbook.addWorksheet('Herramientas y Activos');
      sheetHerramientas.columns = [
        { header: 'HERRAMIENTA / EQUIPO', key: 'nombre', width: 30 },
        { header: 'CATEGORÍA', key: 'categoria', width: 20 },
        { header: 'LOCALIZACIÓN ACTUAL', key: 'ubicacion', width: 28 },
        { header: 'CANTIDAD / UNIDADES', key: 'cantidad', width: 20 },
        { header: 'UNIDAD MEDIDA', key: 'unidad', width: 16 },
        { header: 'ESTADO OPERATIVO', key: 'estado', width: 25 }
      ];

      itemsHerramientas.forEach(h => {
        sheetHerramientas.addRow({
          nombre: h.nombre_insumo.toUpperCase(),
          categoria: (h.categoria || 'HERRAMIENTA').toUpperCase(),
          ubicacion: (h.ubicacion || 'BODEGA PRINCIPAL').toUpperCase(),
          cantidad: parseFloat(h.cantidad_actual) || 1,
          unidad: h.unidad_medida,
          estado: (h.estado_herramienta || 'OPERATIVO').toUpperCase()
        });
      });

      const ultFilaH = sheetHerramientas.rowCount;
      const totalRowH = sheetHerramientas.addRow({
        ubicacion: 'TOTAL UNIDADES:',
        cantidad: { formula: `=SUM(D2:D${ultFilaH})` }
      });

      // Estilizar Hoja 2
      const headH = sheetHerramientas.getRow(1);
      headH.height = 24;
      headH.eachCell(cell => {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF6B21A8' } }; // Color Morado Corporativo
        cell.font = { name: 'Arial', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
      });

      sheetHerramientas.eachRow((row, rNum) => {
        if (rNum === 1 || rNum === ultFilaH + 1) return;
        row.height = 20;
        const cebra = rNum % 2 === 0;
        row.eachCell((cell, colN) => {
          cell.font = { name: 'Arial', size: 9 };
          if (cebra) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF3E8FF' } };
          if ([2, 3, 5, 6].includes(colN)) cell.alignment = { vertical: 'middle', horizontal: 'center' };
          else if (colN === 4) cell.alignment = { vertical: 'middle', horizontal: 'right' };
          else cell.alignment = { vertical: 'middle', horizontal: 'left' };
        });
      });

      totalRowH.height = 22;
      totalRowH.eachCell((cell, colN) => {
        cell.font = { name: 'Arial', size: 10, bold: true, color: { argb: 'FF581C87' } };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE9D5FF' } };
        cell.border = { top: { style: 'thin' }, bottom: { style: 'double' } };
        if (colN === 3) cell.alignment = { vertical: 'middle', horizontal: 'right' };
        if (colN === 4) cell.alignment = { vertical: 'middle', horizontal: 'right' };
      });

      const buffer = await workbook.xlsx.writeBuffer();
      const fechaHoy = new Date().toISOString().split('T')[0];
      const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      saveAs(blob, `INVENTARIO_BODEGA_${fechaHoy}.xlsx`);

      mostrarAlerta("Inventario exportado en 2 hojas con éxito", "exito");
    } catch (err) {
      console.error("Error al exportar Excel:", err);
      mostrarAlerta("Error al generar el archivo de Excel", "error");
    }
  };

  const insumosCriticos = listaInventario.filter(i => i.tipo_item === 'Consumible' && i.aplica_stock !== false && i.cantidad_actual <= 5);

  const insumosFiltrados = listaInventario.filter(i => {
    const coincideTexto = i.nombre_insumo.toLowerCase().includes(busqueda.toLowerCase()) ||
                          i.categoria.toLowerCase().includes(busqueda.toLowerCase()) ||
                          (i.ubicacion && i.ubicacion.toLowerCase().includes(busqueda.toLowerCase()));
    if (filtroTipo === 'Consumible') return coincideTexto && i.tipo_item === 'Consumible';
    if (filtroTipo === 'Herramienta') return coincideTexto && i.tipo_item !== 'Consumible';
    return coincideTexto;
  });

  return (
    <div className="space-y-6 pb-20 text-slate-800">
      
      {/* BANNER DE ALERTAS (SÓLO CONSUMIBLES) */}
      {insumosCriticos.length > 0 && (
        <div className="bg-amber-50 border-l-8 border-amber-500 p-4 rounded-2xl shadow-md">
          <h4 className="font-black text-slate-800 text-xs uppercase tracking-wider mb-2">⚠️ Alertas de Agotamiento de Insumos</h4>
          <div className="flex flex-wrap gap-2">
            {insumosCriticos.map(ins => {
              const estado = obtenerEstadoStock(ins);
              return (
                <span key={ins.id} className={`font-bold text-[11px] px-3 py-1 rounded-lg uppercase ${estado.claseTabla}`}>
                  {ins.nombre_insumo}: <b className="font-black">{ins.cantidad_actual}</b> {ins.unidad_medida}(s) [{estado.etiqueta}]
                </span>
              );
            })}
          </div>
        </div>
      )}

      {/* MENÚ DE SECCIONES */}
      <div className="flex gap-2 border-b pb-2 overflow-x-auto">
        <button onClick={() => setTabActiva('bodega')} className={`px-4 py-2 rounded-xl font-black text-xs uppercase tracking-wider whitespace-nowrap transition-all ${tabActiva === 'bodega' ? 'bg-[#117097] text-white shadow' : 'bg-gray-200 text-gray-600'}`}>📦 Existencias / Ubicaciones</button>
        <button onClick={() => setTabActiva('entrada')} className={`px-4 py-2 rounded-xl font-black text-xs uppercase tracking-wider whitespace-nowrap transition-all ${tabActiva === 'entrada' ? 'bg-[#117097] text-white shadow' : 'bg-gray-200 text-gray-600'}`}>📥 Compras / Entradas Insumos</button>
        <button onClick={() => setTabActiva('salida')} className={`px-4 py-2 rounded-xl font-black text-xs uppercase tracking-wider whitespace-nowrap transition-all ${tabActiva === 'salida' ? 'bg-[#117097] text-white shadow' : 'bg-gray-200 text-gray-600'}`}>🧪 Consumo / Aplicación Dosis</button>
        <button onClick={() => setTabActiva('reubicar')} className={`px-4 py-2 rounded-xl font-black text-xs uppercase tracking-wider whitespace-nowrap transition-all ${tabActiva === 'reubicar' ? 'bg-[#117097] text-white shadow' : 'bg-gray-200 text-gray-600'}`}>🔄 Mover / Asignar Herramientas</button>
      </div>

      {tabActiva === 'bodega' && (
        <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
          
          {esAdmin && (
            <div className="space-y-4">
              {/* REGISTRO DE NUEVO ARTÍCULO */}
              <div className="bg-white p-5 rounded-3xl shadow-xl border border-gray-200 space-y-3">
                <h3 className="font-black text-slate-800 uppercase text-xs italic">➕ Registrar Insumo o Herramienta</h3>
                <form onSubmit={crearInsumoBauche} className="space-y-3">
                  <div>
                    <label className="text-[10px] font-bold text-gray-400 uppercase italic">Tipo de Artículo</label>
                    <select className="w-full border-2 p-2 rounded-xl font-black text-xs bg-sky-50 text-[#117097] outline-none" value={insumoForm.tipo_item} onChange={e => setInsumoForm({ ...insumoForm, tipo_item: e.target.value })}>
                      <option value="Consumible">🧪 CONSUMIBLE (Fertilizante, Veneno, Dosis)</option>
                      <option value="Herramienta">🛠️ HERRAMIENTA / ACTIVO (Motobomba, Manguera, Alambre)</option>
                    </select>
                  </div>

                  <div>
                    <label className="text-[10px] font-bold text-gray-400 uppercase italic">Nombre del Artículo</label>
                    <input type="text" className="w-full border-2 p-2 rounded-xl font-bold bg-gray-50 uppercase text-xs outline-none focus:border-[#117097] focus:bg-white" value={insumoForm.nombre_insumo} onChange={e => setInsumoForm({...insumoForm, nombre_insumo: e.target.value})} placeholder="Ej: Triple 15 / Motobomba Honda" required />
                  </div>

                  <div className="grid grid-cols-2 gap-2">
                    <div>
                      <label className="text-[10px] font-bold text-gray-400 uppercase italic">Categoría</label>
                      <select className="w-full border-2 p-2 rounded-xl font-bold bg-white text-xs outline-none" value={insumoForm.categoria} onChange={e => setInsumoForm({...insumoForm, categoria: e.target.value})}>
                        {categorias.map(cat => <option key={cat} value={cat}>{cat}</option>)}
                      </select>
                    </div>
                    <div>
                      <label className="text-[10px] font-bold text-gray-400 uppercase italic">Unidad Medida</label>
                      <select className="w-full border-2 p-2 rounded-xl font-bold bg-white text-xs outline-none" value={insumoForm.unidad_medida} onChange={e => setInsumoForm({...insumoForm, unidad_medida: e.target.value})}>
                        {unidades.map(un => <option key={un} value={un}>{un}</option>)}
                      </select>
                    </div>
                  </div>

                  <div>
                    <label className="text-[10px] font-bold text-gray-400 uppercase italic">Ubicación / Localización Física</label>
                    <select className="w-full border-2 p-2 rounded-xl font-bold bg-amber-50/40 text-amber-900 text-xs outline-none" value={insumoForm.ubicacion} onChange={e => setInsumoForm({...insumoForm, ubicacion: e.target.value})}>
                      <optgroup label="🏢 Bodegas & Áreas">
                        {listaUbicaciones.map(u => <option key={u.id} value={u.nombre}>{u.nombre}</option>)}
                      </optgroup>
                      <optgroup label="🌱 Invernaderos / Lotes">
                        {datosInvernaderos?.map(inv => <option key={inv.id} value={inv.nombre?.toUpperCase()}>{inv.nombre?.toUpperCase()}</option>)}
                      </optgroup>
                    </select>
                  </div>

                  {insumoForm.tipo_item === 'Consumible' ? (
                    <div className="bg-slate-50 p-3 rounded-xl border border-slate-200 space-y-2">
                      <label className="flex items-center gap-2 cursor-pointer">
                        <input type="checkbox" className="w-4 h-4 accent-[#117097]" checked={insumoForm.aplica_stock} onChange={e => setInsumoForm({ ...insumoForm, aplica_stock: e.target.checked })} />
                        <span className="text-[11px] font-black text-slate-700 uppercase">¿Controlar Stock Mínimo?</span>
                      </label>

                      {insumoForm.aplica_stock && (
                        <div>
                          <label className="text-[10px] font-bold text-gray-400 uppercase italic">Stock Mínimo Deseado</label>
                          <input type="number" className="w-full border-2 p-2 rounded-xl font-bold text-xs bg-white" value={insumoForm.stock_minimo} onChange={e => setInsumoForm({...insumoForm, stock_minimo: e.target.value})} required />
                        </div>
                      )}
                    </div>
                  ) : (
                    <div>
                      <label className="text-[10px] font-bold text-gray-400 uppercase italic">Estado Inicial</label>
                      <select className="w-full border-2 p-2 rounded-xl font-bold bg-white text-xs outline-none" value={insumoForm.estado_herramienta} onChange={e => setInsumoForm({...insumoForm, estado_herramienta: e.target.value})}>
                        {estadosHerramientaOptions.map(est => <option key={est} value={est}>{est}</option>)}
                      </select>
                    </div>
                  )}

                  <button type="submit" className="w-full py-3 bg-[#117097] text-white font-black rounded-xl uppercase text-xs tracking-widest hover:bg-[#0a4c68] transition-colors shadow-md">💾 Guardar Registro</button>
                </form>
              </div>

              {/* CREAR NUEVAS UBICACIONES */}
              <div className="bg-slate-50 p-4 rounded-2xl border-2 border-slate-200 space-y-2">
                <p className="text-[10px] font-black text-slate-600 uppercase italic">📍 Nueva Ubicación / Bodega Física</p>
                <form onSubmit={agregarNuevaUbicacion} className="flex gap-2">
                  <input type="text" className="w-full border p-2 bg-white rounded-xl text-xs font-bold uppercase outline-none" value={nuevaUbicacionTexto} onChange={e => setNuevaUbicacionTexto(e.target.value)} placeholder="Ej: Cuarto de Herramientas" required />
                  <button type="submit" className="px-3 bg-emerald-600 text-white font-black rounded-xl text-xs uppercase">+</button>
                </form>
              </div>
            </div>
          )}

          {/* TABLA DE BODEGA CON BOTÓN EXPORTAR EXCEL */}
          <div className={`${esAdmin ? 'lg:col-span-2' : 'lg:col-span-3'} bg-white rounded-3xl shadow-xl overflow-hidden border border-gray-200 flex flex-col`}>
            
            {/* CABECERA CON FILTROS Y BOTÓN EXCEL */}
            <div className="p-4 bg-[#117097] text-white font-black text-xs uppercase tracking-wider flex flex-col xl:flex-row xl:items-center justify-between gap-3">
              
              <div className="flex items-center gap-2 flex-wrap">
                <button onClick={() => setFiltroTipo('TODOS')} className={`px-2.5 py-1 rounded-lg text-[10px] uppercase font-black transition-all ${filtroTipo === 'TODOS' ? 'bg-white text-[#117097]' : 'bg-[#0a4c68] text-sky-100'}`}>Todos ({listaInventario.length})</button>
                <button onClick={() => setFiltroTipo('Consumible')} className={`px-2.5 py-1 rounded-lg text-[10px] uppercase font-black transition-all ${filtroTipo === 'Consumible' ? 'bg-white text-[#117097]' : 'bg-[#0a4c68] text-sky-100'}`}>🧪 Consumibles</button>
                <button onClick={() => setFiltroTipo('Herramienta')} className={`px-2.5 py-1 rounded-lg text-[10px] uppercase font-black transition-all ${filtroTipo === 'Herramienta' ? 'bg-white text-[#117097]' : 'bg-[#0a4c68] text-sky-100'}`}>🛠️ Herramientas / Activos</button>
              </div>

              <div className="flex items-center gap-2 flex-wrap justify-end">
                {/* ⚡ BOTÓN PARA EXPORTAR A EXCEL EN 2 HOJAS */}
                <button
                  onClick={exportarInventarioAExcel}
                  className="px-3 py-1.5 bg-emerald-600 hover:bg-emerald-700 text-white text-[10px] font-black uppercase rounded-lg shadow transition-colors flex items-center gap-1"
                >
                  📊 EXPORTAR A EXCEL
                </button>

                <input 
                  type="text" 
                  placeholder="🔍 Buscar por nombre o ubicación..." 
                  className="px-3 py-1.5 text-xs rounded-xl text-slate-800 outline-none font-bold placeholder-gray-400 min-w-[200px]" 
                  value={busqueda} 
                  onChange={e => setBusqueda(e.target.value)} 
                />
              </div>

            </div>

            <div className="overflow-x-auto min-h-[350px]">
              <table className="w-full text-left border-collapse text-[11px]">
                <thead>
                  <tr className="bg-slate-100 text-[10px] font-black uppercase text-slate-600 border-b">
                    <th className="p-3">Artículo / Nombre</th>
                    <th className="p-3 text-center">Tipo</th>
                    <th className="p-3 text-center">📍 Localización Actual</th>
                    <th className="p-3 text-right">Cant. Actual</th>
                    <th className="p-3 text-center">Unidad</th>
                    <th className="p-3 text-center">Estado / Alerta</th>
                    {esAdmin && <th className="p-3 text-center">Acciones</th>}
                  </tr>
                </thead>
                <tbody className="divide-y font-bold text-slate-700">
                  {insumosFiltrados.length === 0 ? (
                    <tr><td colSpan={esAdmin ? 7 : 6} className="p-8 text-center text-gray-400 italic font-bold">No hay artículos coincidentes</td></tr>
                  ) : (
                    insumosFiltrados.map((ins, idx) => {
                      const estado = obtenerEstadoStock(ins);
                      const editandoEste = idEditando === ins.id;

                      return (
                        <tr key={ins.id} className={idx % 2 === 0 ? 'bg-white' : 'bg-slate-50/60 hover:bg-sky-50 transition-colors'}>
                          <td className="p-3 font-black text-slate-900 uppercase">
                            {editandoEste ? (
                              <input type="text" className="border-2 px-2 py-1 rounded-lg w-full bg-white text-xs uppercase" value={filaEditable.nombre_insumo} onChange={e => setFilaEditable({...filaEditable, nombre_insumo: e.target.value})} />
                            ) : (
                              <div>
                                <p className="font-black text-slate-900">{ins.nombre_insumo}</p>
                                <span className="text-[9px] text-gray-400 font-bold">{ins.categoria}</span>
                              </div>
                            )}
                          </td>

                          <td className="p-3 text-center">
                            <span className={`px-2 py-0.5 rounded text-[8px] font-black uppercase ${ins.tipo_item === 'Consumible' ? 'bg-sky-100 text-[#117097]' : 'bg-purple-100 text-purple-800'}`}>
                              {ins.tipo_item === 'Consumible' ? '🧪 Consumible' : '🛠️ Activo'}
                            </span>
                          </td>

                          <td className="p-3 text-center">
                            <span className="inline-block bg-amber-100/80 text-amber-900 px-2 py-0.5 rounded-lg text-[9px] font-black uppercase">
                              📍 {ins.ubicacion || 'BODEGA PRINCIPAL'}
                            </span>
                          </td>

                          <td className={`p-3 text-right font-black text-xs ${estado.claseTexto}`}>
                            {editandoEste ? (
                              <input type="number" step="any" className="border-2 px-2 py-1 rounded-lg w-20 text-right" value={filaEditable.cantidad_actual} onChange={e => setFilaEditable({...filaEditable, cantidad_actual: e.target.value})} />
                            ) : (
                              ins.cantidad_actual
                            )}
                          </td>

                          <td className="p-3 text-center text-gray-400">{ins.unidad_medida}</td>

                          <td className="p-3 text-center">
                            <span className={`px-2 py-0.5 rounded-full text-[8px] font-black uppercase ${estado.claseTabla}`}>
                              {estado.etiqueta}
                            </span>
                          </td>

                          {esAdmin && (
                            <td className="p-3 text-center">
                              {editandoEste ? (
                                <div className="flex gap-1 justify-center">
                                  <button onClick={() => guardarEdicionFila(ins.id)} className="bg-emerald-600 text-white px-2 py-1 rounded text-[9px] font-black">💾</button>
                                  <button onClick={() => setIdEditando(null)} className="bg-gray-400 text-white px-2 py-1 rounded text-[9px] font-black">❌</button>
                                </div>
                              ) : (
                                <div className="flex gap-1 justify-center">
                                  <button onClick={() => iniciarEdicion(ins)} className="p-1 bg-amber-100 text-amber-700 rounded border hover:bg-amber-600 hover:text-white text-xs">✏️</button>
                                  <button onClick={() => eliminarInsumoCompleto(ins.id, ins.nombre_insumo)} className="p-1 bg-red-100 text-red-700 rounded border hover:bg-red-600 hover:text-white text-xs">🗑️</button>
                                </div>
                              )}
                            </td>
                          )}
                        </tr>
                      );
                    })
                  )}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      )}

      {/* COMPRAS */}
      {tabActiva === 'entrada' && (
        <div className="bg-white p-6 rounded-3xl shadow-xl border-t-8 border-[#117097] max-w-2xl mx-auto">
          <h3 className="font-black text-[#117097] uppercase text-sm mb-4 italic">📥 Compra e Ingreso de Insumos Consumibles</h3>
          <form onSubmit={registrarEntradaBodega} className="space-y-4">
            <div>
              <label className="text-[10px] font-bold text-gray-400 uppercase italic">Seleccione Insumo Consumible</label>
              <select className="w-full border-2 p-3 rounded-xl font-bold bg-white text-xs outline-none focus:border-[#117097]" value={entradaForm.insumo_id} onChange={e => setEntradaForm({...entradaForm, insumo_id: e.target.value})} required>
                <option value="">Seleccione insumo...</option>
                {listaInventario.filter(i => i.tipo_item === 'Consumible').map(i => (
                  <option key={i.id} value={i.id}>{i.nombre_insumo} ({i.unidad_medida}) — Stock Actual: {i.cantidad_actual}</option>
                ))}
              </select>
            </div>
            <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
              <div>
                <label className="text-[10px] font-bold text-gray-400 uppercase italic">Cantidad Ingresada</label>
                <input type="number" step="any" className="w-full border-2 p-3 rounded-xl font-black text-lg text-emerald-700 outline-none focus:border-[#117097]" value={entradaForm.cantidad_ingresada} onChange={e => setEntradaForm({...entradaForm, cantidad_ingresada: e.target.value})} required />
              </div>
              <div>
                <label className="text-[10px] font-bold text-gray-400 uppercase italic">P. Unitario de Compra (COP)</label>
                <input type="number" className="w-full border-2 p-3 rounded-xl font-black text-lg text-slate-700 outline-none focus:border-[#117097]" value={entradaForm.precio_unitario_compra} onChange={e => setEntradaForm({...entradaForm, precio_unitario_compra: e.target.value})} placeholder="Opcional" />
              </div>
            </div>
            <div>
              <label className="text-[10px] font-bold text-gray-400 uppercase italic">N° Comprobante / Factura</label>
              <input type="text" className="w-full border-2 p-3 rounded-xl font-bold outline-none focus:border-[#117097]" value={entradaForm.numero_factura_comprobante} onChange={e => setEntradaForm({...entradaForm, numero_factura_comprobante: e.target.value})} placeholder="Ej: FAC-4589" />
            </div>
            <div>
              <label className="text-[10px] font-bold text-gray-400 uppercase italic">Observaciones / Proveedor</label>
              <textarea className="w-full border-2 p-3 rounded-xl font-bold h-20 bg-gray-50 uppercase focus:bg-white text-xs outline-none focus:border-[#117097]" value={entradaForm.observaciones} onChange={e => setEntradaForm({...entradaForm, observaciones: e.target.value})} placeholder="Ej: Comprado en Almacén El Ganadero" />
            </div>
            <button type="submit" className="w-full py-4 bg-[#117097] hover:bg-[#0a4c68] transition-colors text-white font-black rounded-2xl uppercase tracking-widest text-xs shadow-lg">📈 Confirmar Compra e Incrementar Inventario</button>
          </form>
        </div>
      )}

      {/* CONSUMO / DOSIS */}
      {tabActiva === 'salida' && (
        <div className="bg-white p-6 rounded-3xl shadow-xl border-t-8 border-[#117097] max-w-2xl mx-auto">
          <h3 className="font-black text-[#117097] uppercase text-sm mb-4 italic">🧪 Aplicación de Dosis / Consumo en Cultivo</h3>
          <form onSubmit={registrarSalidaInvernadero} className="space-y-4">
            <div>
              <label className="text-[10px] font-bold text-gray-400 uppercase italic">Insumo Consumible a Aplicar</label>
              <select className="w-full border-2 p-3 rounded-xl font-bold bg-white text-xs outline-none focus:border-[#117097]" value={salidaForm.insumo_id} onChange={e => setSalidaForm({...salidaForm, insumo_id: e.target.value})} required>
                <option value="">Seleccione insumo...</option>
                {listaInventario.filter(i => i.tipo_item === 'Consumible').map(i => (
                  <option key={i.id} value={i.id}>{i.nombre_insumo} — Disponible: {i.cantidad_actual} {i.unidad_medida}(s)</option>
                ))}
              </select>
            </div>
            <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
              <div>
                <label className="text-[10px] font-bold text-gray-400 uppercase italic">Cantidad Consumida / Dosis</label>
                <input type="number" step="any" className="w-full border-2 p-3 rounded-xl font-black text-lg text-red-600 outline-none focus:border-[#117097]" value={salidaForm.cantidad_retirada} onChange={e => setSalidaForm({...salidaForm, cantidad_retirada: e.target.value})} required />
              </div>
              <div>
                <label className="text-[10px] font-bold text-gray-400 uppercase italic">Invernadero / Cultivo Destino</label>
                <select className="w-full border-2 p-3 rounded-xl font-bold bg-white text-xs outline-none focus:border-[#117097]" value={salidaForm.invernadero_id} onChange={e => setSalidaForm({...salidaForm, invernadero_id: e.target.value})} required>
                  <option value="">Seleccione Destino...</option>
                  {datosInvernaderos?.map(inv => (
                    <option key={inv.id} value={inv.id}>{inv.nombre?.toUpperCase() || inv.nombre_invernadero?.toUpperCase()}</option>
                  ))}
                </select>
              </div>
            </div>

            <div>
              <label className="text-[10px] font-bold text-gray-400 uppercase italic">Responsable Aplicador (Desde Nómina)</label>
              <select className="w-full border-2 p-3 rounded-xl font-bold uppercase bg-white text-xs outline-none focus:border-[#117097]" value={salidaForm.responsable} onChange={e => setSalidaForm({ ...salidaForm, responsable: e.target.value })} required>
                <option value="">Seleccione operario...</option>
                {listaTrabajadores.map(trab => (
                  <option key={trab.id} value={trab.nombre_completo}>{trab.nombre_completo}</option>
                ))}
              </select>
            </div>

            <div>
              <label className="text-[10px] font-bold text-gray-400 uppercase italic">Nota de Aplicación / Dosis / Plaga</label>
              <textarea className="w-full border-2 p-3 rounded-xl font-bold h-20 bg-gray-50 uppercase focus:bg-white text-xs outline-none focus:border-[#117097]" value={salidaForm.nota_uso} onChange={e => setSalidaForm({...salidaForm, nota_uso: e.target.value})} placeholder="Ej: Control de roya con bomba de espalda" />
            </div>
            <button type="submit" className="w-full py-4 bg-[#117097] hover:bg-[#0a4c68] transition-colors text-white font-black rounded-2xl uppercase tracking-widest text-xs shadow-lg">📉 Confirmar Aplicación y Descontar de Bodega</button>
          </form>
        </div>
      )}

      {/* REUBICAR HERRAMIENTAS */}
      {tabActiva === 'reubicar' && (
        <div className="bg-white p-6 rounded-3xl shadow-xl border-t-8 border-purple-800 max-w-2xl mx-auto">
          <h3 className="font-black text-purple-900 uppercase text-sm mb-4 italic">🔄 Traslado y Asignación de Herramientas / Equipos</h3>
          <form onSubmit={reubicarHerramienta} className="space-y-4">
            <div>
              <label className="text-[10px] font-bold text-gray-400 uppercase italic">Seleccione Herramienta o Equipo</label>
              <select className="w-full border-2 p-3 rounded-xl font-bold bg-white text-xs outline-none focus:border-purple-800" value={reubicacionForm.herramienta_id} onChange={e => setReubicacionForm({ ...reubicacionForm, herramienta_id: e.target.value })} required>
                <option value="">Seleccione activo...</option>
                {listaInventario.filter(i => i.tipo_item !== 'Consumible').map(i => (
                  <option key={i.id} value={i.id}>{i.nombre_insumo} — Ubicación Actual: {i.ubicacion || 'BODEGA PRINCIPAL'}</option>
                ))}
              </select>
            </div>

            <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
              <div>
                <label className="text-[10px] font-bold text-gray-400 uppercase italic">Nueva Ubicación Destino</label>
                <select className="w-full border-2 p-3 rounded-xl font-bold bg-amber-50/40 text-amber-900 text-xs outline-none focus:border-purple-800" value={reubicacionForm.nueva_ubicacion} onChange={e => setReubicacionForm({ ...reubicacionForm, nueva_ubicacion: e.target.value })} required>
                  <option value="">Seleccione nueva ubicación...</option>
                  <optgroup label="🌱 Invernaderos / Cultivos">
                    {datosInvernaderos?.map(inv => <option key={inv.id} value={`INVERNADERO ${inv.nombre?.toUpperCase()}`}>{inv.nombre?.toUpperCase()}</option>)}
                  </optgroup>
                  <optgroup label="🏢 Bodegas & Zonas">
                    {listaUbicaciones.map(u => <option key={u.id} value={u.nombre}>{u.nombre}</option>)}
                  </optgroup>
                </select>
              </div>

              <div>
                <label className="text-[10px] font-bold text-gray-400 uppercase italic">Operario Responsable / Asignado</label>
                <select className="w-full border-2 p-3 rounded-xl font-bold uppercase bg-white text-xs outline-none focus:border-purple-800" value={reubicacionForm.responsable} onChange={e => setReubicacionForm({ ...reubicacionForm, responsable: e.target.value })}>
                  <option value="">Sin responsable fijo...</option>
                  {listaTrabajadores.map(trab => (
                    <option key={trab.id} value={trab.nombre_completo}>{trab.nombre_completo}</option>
                  ))}
                </select>
              </div>
            </div>

            <div>
              <label className="text-[10px] font-bold text-gray-400 uppercase italic">Observaciones del Traslado / Estado</label>
              <textarea className="w-full border-2 p-3 rounded-xl font-bold h-20 bg-gray-50 uppercase focus:bg-white text-xs outline-none focus:border-purple-800" value={reubicacionForm.observaciones} onChange={e => setReubicacionForm({ ...reubicacionForm, observaciones: e.target.value })} placeholder="Ej: Se entrega con tanque lleno de gasolina" />
            </div>

            <button type="submit" className="w-full py-4 bg-purple-800 hover:bg-purple-900 transition-colors text-white font-black rounded-2xl uppercase tracking-widest text-xs shadow-lg">🔄 Actualizar Localización de Herramienta</button>
          </form>
        </div>
      )}

    </div>
  );
}