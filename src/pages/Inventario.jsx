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

  // ESTADO MODERNO PARA MODALES (Reemplaza las pestañas)
  const [modalActivo, setModalActivo] = useState(null); // null, 'nuevo', 'entrada', 'salida', 'reubicar'
  
  const [filtroTipo, setFiltroTipo] = useState('TODOS'); 
  const [cargando, setCargando] = useState(false);
  const [busqueda, setBusqueda] = useState('');

  // Edición directa extendida
  const [idEditando, setIdEditando] = useState(null);
  const [filaEditable, setFilaEditable] = useState({ 
    nombre_insumo: '', categoria: '', unidad_medida: '', cantidad_actual: 0, 
    stock_minimo: 0, aplica_stock: true, tipo_item: 'Consumible', ubicacion: 'BODEGA PRINCIPAL', estado_herramienta: 'Operativo'
  });

  // Formularios
  const [insumoForm, setInsumoForm] = useState({ 
    nombre_insumo: '', tipo_item: 'Consumible', categoria: 'Fertilizante', 
    unidad_medida: 'Bulto', cantidad_inicial: 1, stock_minimo: 3, aplica_stock: true, ubicacion: 'BODEGA PRINCIPAL', estado_herramienta: 'Operativo'
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

  const categorias = ['Fertilizante', 'Fungicida', 'Semilla', 'Herramienta', 'Maquinaria', 'Canastilla', 'Plastico', 'Accesorio', 'Manguera', 'Otros'];
  const unidades = ['Bulto', 'Kilo', 'Litro', 'Metro', 'Unidad', 'Caja','Libra','Gramos', 'Galón'];
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

  const agregarNuevaUbicacion = async () => {
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
      const cantIngresada = parseFloat(insumoForm.cantidad_inicial) >= 0 ? parseFloat(insumoForm.cantidad_inicial) : (esConsumible ? 0 : 1);

      const { error } = await supabase.from('inventario').insert([{
        nombre_insumo: insumoForm.nombre_insumo.toUpperCase().trim(),
        tipo_item: insumoForm.tipo_item,
        categoria: insumoForm.categoria,
        unidad_medida: insumoForm.unidad_medida,
        stock_minimo: esConsumible && insumoForm.aplica_stock ? (parseFloat(insumoForm.stock_minimo) || 0) : 0,
        cantidad_actual: cantIngresada,
        aplica_stock: esConsumible ? insumoForm.aplica_stock : false,
        ubicacion: insumoForm.ubicacion.toUpperCase().trim(),
        estado_herramienta: insumoForm.estado_herramienta
      }]);

      if (error) throw error;
      mostrarAlerta("Artículo registrado en bodega", "exito");
      setInsumoForm({ nombre_insumo: '', tipo_item: 'Consumible', categoria: 'Fertilizante', unidad_medida: 'Bulto', cantidad_inicial: 1, stock_minimo: 3, aplica_stock: true, ubicacion: 'BODEGA PRINCIPAL', estado_herramienta: 'Operativo' });
      setModalActivo(null); 
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
      setModalActivo(null); 
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
      setModalActivo(null); 
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
      setModalActivo(null); 
      cargarInventario();
    } catch (err) { mostrarAlerta("Error al reubicar la herramienta", "error"); }
  };

  const iniciarEdicion = (ins) => {
    if (!esAdmin) return;
    setIdEditando(ins.id);
    setFilaEditable({
      nombre_insumo: ins.nombre_insumo || '',
      categoria: ins.categoria || 'Fertilizante',
      unidad_medida: ins.unidad_medida || 'Unidad',
      cantidad_actual: ins.cantidad_actual ?? 0,
      stock_minimo: ins.stock_minimo ?? 0,
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
      return { nivel: 'herramienta', etiqueta: ins.estado_herramienta || 'Activo', claseTabla: 'bg-slate-100 dark:bg-slate-700 text-slate-800 dark:text-slate-300 border border-slate-300 dark:border-slate-600', claseTexto: 'text-[#117097] dark:text-sky-400' };
    }
    const cant = parseFloat(ins.cantidad_actual) || 0;
    if (cant <= 1) {
      return { nivel: 'rojo', etiqueta: 'CRÍTICO (≤1)', claseTabla: 'bg-red-500 dark:bg-red-700 text-white animate-pulse shadow-sm', claseTexto: 'text-red-600 dark:text-red-400 font-black' };
    } else if (cant >= 2 && cant <= 3) {
      return { nivel: 'amarillo', etiqueta: 'ALERTA (2-3)', claseTabla: 'bg-amber-400 dark:bg-amber-600 text-amber-900 dark:text-amber-100 shadow-sm', claseTexto: 'text-amber-600 dark:text-amber-400 font-black' };
    }
    return { nivel: 'verde', etiqueta: 'DISPONIBLE', claseTabla: 'bg-emerald-100 dark:bg-emerald-950/80 text-emerald-800 dark:text-emerald-400 border border-emerald-200 dark:border-emerald-800', claseTexto: 'text-emerald-700 dark:text-emerald-400' };
  };

  const exportarInventarioAExcel = async () => {
    if (!listaInventario || listaInventario.length === 0) {
      mostrarAlerta("No hay artículos en bodega para exportar", "error");
      return;
    }
    try {
      const workbook = new ExcelJS.Workbook();
      const itemsConsumibles = listaInventario.filter(i => i.tipo_item === 'Consumible');
      const itemsHerramientas = listaInventario.filter(i => i.tipo_item !== 'Consumible');

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
        row.eachCell((cell, colN) => {
          cell.font = { name: 'Arial', size: 9 };
          if (rNum % 2 === 0) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEBF5FB' } };
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
        if (colN === 3 || colN === 4) cell.alignment = { vertical: 'middle', horizontal: 'right' };
      });

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

      const headH = sheetHerramientas.getRow(1);
      headH.height = 24;
      headH.eachCell(cell => {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF6B21A8' } };
        cell.font = { name: 'Arial', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
      });

      sheetHerramientas.eachRow((row, rNum) => {
        if (rNum === 1 || rNum === ultFilaH + 1) return;
        row.height = 20;
        row.eachCell((cell, colN) => {
          cell.font = { name: 'Arial', size: 9 };
          if (rNum % 2 === 0) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF3E8FF' } };
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
        if (colN === 3 || colN === 4) cell.alignment = { vertical: 'middle', horizontal: 'right' };
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

  const insumosCriticos = listaInventario.filter(i => i.tipo_item === 'Consumible' && i.aplica_stock !== false && parseFloat(i.cantidad_actual) <= 3);

  const insumosFiltrados = listaInventario.filter(i => {
    const coincideTexto = i.nombre_insumo.toLowerCase().includes(busqueda.toLowerCase()) ||
                          i.categoria.toLowerCase().includes(busqueda.toLowerCase()) ||
                          (i.ubicacion && i.ubicacion.toLowerCase().includes(busqueda.toLowerCase()));
    if (filtroTipo === 'Consumible') return coincideTexto && i.tipo_item === 'Consumible';
    if (filtroTipo === 'Herramienta') return coincideTexto && i.tipo_item !== 'Consumible';
    return coincideTexto;
  });

  return (
    <div className="space-y-6 pb-20 text-slate-800 dark:text-slate-200 font-sans transition-colors duration-300">
      
      {/* 🚀 CABECERA PRINCIPAL MODERNA */}
      <div className="bg-white dark:bg-slate-800 p-6 rounded-3xl shadow-xl border border-slate-200 dark:border-slate-700 flex flex-col xl:flex-row justify-between items-center gap-4 transition-colors duration-300">
        <div className="space-y-1 w-full xl:w-auto text-center xl:text-left">
          <div className="flex items-center justify-center xl:justify-start gap-3">
            <span className="text-2xl p-2 bg-[#117097]/10 dark:bg-sky-500/20 rounded-xl text-[#117097] dark:text-sky-400">📦</span>
            <h2 className="text-xl font-black tracking-tight text-slate-900 dark:text-white uppercase">Inventario de Bodega</h2>
          </div>
          <p className="text-xs text-slate-500 dark:text-slate-400 font-medium">Control de consumibles, herramientas y existencias físicas.</p>
        </div>
        
        <div className="flex items-center gap-2 flex-wrap justify-center xl:justify-end w-full xl:w-auto">
          <button onClick={exportarInventarioAExcel} className="px-4 py-2.5 bg-emerald-600 hover:bg-emerald-700 text-white font-black text-xs uppercase tracking-wider rounded-xl shadow-md transition-all flex items-center gap-1.5 cursor-pointer">
            <span>📊</span> Excel
          </button>
          {esAdmin && (
             <button onClick={() => setModalActivo('nuevo')} className="px-4 py-2.5 bg-[#117097] hover:bg-[#0a4c68] text-white font-black text-xs uppercase tracking-wider rounded-xl shadow-md transition-all flex items-center gap-1.5 cursor-pointer">
               <span>➕</span> Nuevo Artículo
             </button>
          )}
          <button onClick={() => setModalActivo('entrada')} className="px-4 py-2.5 bg-emerald-700 hover:bg-emerald-800 text-white font-black text-xs uppercase tracking-wider rounded-xl shadow-md transition-all flex items-center gap-1.5 cursor-pointer">
            <span>📥</span> Ingreso
          </button>
          <button onClick={() => setModalActivo('salida')} className="px-4 py-2.5 bg-rose-600 hover:bg-rose-700 text-white font-black text-xs uppercase tracking-wider rounded-xl shadow-md transition-all flex items-center gap-1.5 cursor-pointer">
            <span>🧪</span> Consumo
          </button>
          <button onClick={() => setModalActivo('reubicar')} className="px-4 py-2.5 bg-purple-700 hover:bg-purple-800 text-white font-black text-xs uppercase tracking-wider rounded-xl shadow-md transition-all flex items-center gap-1.5 cursor-pointer">
            <span>🔄</span> Reubicar
          </button>
        </div>
      </div>

      {/* BANNER DE ALERTAS */}
      {insumosCriticos.length > 0 && (
        <div className="bg-amber-50 dark:bg-amber-950/40 border border-amber-200 dark:border-amber-800/60 p-4 rounded-2xl shadow-sm flex items-start gap-3 transition-colors duration-300">
          <span className="text-2xl animate-pulse">⚠️</span>
          <div>
            <h4 className="font-black text-amber-900 dark:text-amber-400 text-xs uppercase tracking-wider mb-2">Alertas de Agotamiento de Insumos</h4>
            <div className="flex flex-wrap gap-2">
              {insumosCriticos.map(ins => {
                const estado = obtenerEstadoStock(ins);
                return (
                  <span key={ins.id} className={`font-bold text-[10px] px-2.5 py-1 rounded-md uppercase ${estado.claseTabla}`}>
                    {ins.nombre_insumo}: <b className="font-black">{ins.cantidad_actual}</b> {ins.unidad_medida}(s) [{estado.etiqueta}]
                  </span>
                );
              })}
            </div>
          </div>
        </div>
      )}

      {/* 🔍 BARRA DE BÚSQUEDA Y FILTROS */}
      <div className="bg-white dark:bg-slate-800 p-4 rounded-2xl shadow-md border border-slate-200 dark:border-slate-700 flex flex-col md:flex-row justify-between items-center gap-4 transition-colors duration-300">
        <div className="flex items-center gap-2 overflow-x-auto w-full md:w-auto pb-1 md:pb-0">
          <button onClick={() => setFiltroTipo('TODOS')} className={`px-4 py-2 rounded-xl text-[10px] uppercase font-black transition-all whitespace-nowrap cursor-pointer ${filtroTipo === 'TODOS' ? 'bg-[#117097] text-white shadow' : 'bg-slate-100 dark:bg-slate-700 text-slate-500 dark:text-slate-300 hover:bg-slate-200 dark:hover:bg-slate-600'}`}>
            Todos ({listaInventario.length})
          </button>
          <button onClick={() => setFiltroTipo('Consumible')} className={`px-4 py-2 rounded-xl text-[10px] uppercase font-black transition-all whitespace-nowrap cursor-pointer ${filtroTipo === 'Consumible' ? 'bg-emerald-600 text-white shadow' : 'bg-slate-100 dark:bg-slate-700 text-slate-500 dark:text-slate-300 hover:bg-slate-200 dark:hover:bg-slate-600'}`}>
            🧪 Consumibles
          </button>
          <button onClick={() => setFiltroTipo('Herramienta')} className={`px-4 py-2 rounded-xl text-[10px] uppercase font-black transition-all whitespace-nowrap cursor-pointer ${filtroTipo === 'Herramienta' ? 'bg-purple-600 text-white shadow' : 'bg-slate-100 dark:bg-slate-700 text-slate-500 dark:text-slate-300 hover:bg-slate-200 dark:hover:bg-slate-600'}`}>
            🛠️ Activos Fijos
          </button>
        </div>
        
        <div className="relative w-full md:w-96">
          <span className="absolute inset-y-0 left-0 flex items-center pl-3 text-slate-400">🔍</span>
          <input 
            type="text" 
            placeholder="Buscar por nombre, categoría o ubicación..." 
            className="w-full bg-slate-50 dark:bg-slate-700/50 border border-slate-200 dark:border-slate-600 text-xs text-slate-800 dark:text-white rounded-xl pl-10 pr-4 py-2 outline-none focus:border-[#117097] dark:focus:border-sky-500 font-bold placeholder-slate-400"
            value={busqueda}
            onChange={e => setBusqueda(e.target.value)}
          />
        </div>
      </div>

      {/* 📊 TABLA DE INVENTARIO COMPRIMIDA */}
      <div className="bg-white dark:bg-slate-800 rounded-3xl shadow-xl overflow-hidden border border-gray-200 dark:border-slate-700 transition-colors duration-300">
        <div className="overflow-x-auto min-h-[400px]">
          <table className="w-full text-left border-collapse text-xs">
            <thead>
              <tr className="bg-slate-100 dark:bg-slate-700/80 text-[10px] font-black uppercase text-slate-600 dark:text-slate-300 border-b border-slate-200 dark:border-slate-700 sticky top-0 z-10">
                <th className="py-2.5 px-3">Artículo / Nombre</th>
                <th className="py-2.5 px-3 text-center">Tipo</th>
                <th className="py-2.5 px-3 text-center">📍 Localización Actual</th>
                <th className="py-2.5 px-3 text-right">Cant. Actual</th>
                <th className="py-2.5 px-3 text-center">Unidad</th>
                <th className="py-2.5 px-3 text-center">Estado / Alerta</th>
                {esAdmin && <th className="py-2.5 px-3 text-center">Acciones</th>}
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100 dark:divide-slate-700/60 font-bold text-slate-700 dark:text-slate-300">
              {insumosFiltrados.length === 0 ? (
                <tr><td colSpan={esAdmin ? 7 : 6} className="p-6 text-center text-gray-400 dark:text-slate-500 italic font-bold">No hay artículos coincidentes en el inventario.</td></tr>
              ) : (
                insumosFiltrados.map((ins, idx) => {
                  const estado = obtenerEstadoStock(ins);
                  const editandoEste = idEditando === ins.id;

                  return (
                    <tr key={ins.id} className={`${idx % 2 === 0 ? 'bg-white dark:bg-slate-800' : 'bg-slate-50/50 dark:bg-slate-800/60'} hover:bg-sky-50/50 dark:hover:bg-slate-700/50 transition-colors border-l-8 ${ins.tipo_item === 'Consumible' ? 'border-emerald-500' : 'border-purple-500'}`}>
                      
                      <td className="py-2 px-3 font-black text-slate-900 dark:text-white uppercase max-w-[300px] whitespace-normal break-words">
                        {editandoEste ? (
                          <div className="space-y-1">
                            <input type="text" className="border border-slate-300 dark:border-slate-600 p-1.5 rounded-lg w-full bg-white dark:bg-slate-900 text-xs font-black uppercase outline-none focus:border-[#117097] text-slate-800 dark:text-white" value={filaEditable.nombre_insumo} onChange={e => setFilaEditable({...filaEditable, nombre_insumo: e.target.value})} placeholder="Nombre..." />
                            <select className="border border-slate-300 dark:border-slate-600 p-1.5 rounded-lg w-full bg-white dark:bg-slate-900 text-[10px] font-bold outline-none text-slate-800 dark:text-white" value={filaEditable.categoria} onChange={e => setFilaEditable({...filaEditable, categoria: e.target.value})}>
                              {categorias.map(cat => <option key={cat} value={cat}>{cat}</option>)}
                            </select>
                          </div>
                        ) : (
                          <div>
                            <p className="font-black text-slate-900 dark:text-white leading-snug">{ins.nombre_insumo}</p>
                            <span className="text-[9px] text-[#117097] dark:text-sky-400 font-bold">📌 {ins.categoria}</span>
                          </div>
                        )}
                      </td>

                      <td className="py-2 px-3 text-center whitespace-nowrap">
                        {editandoEste ? (
                          <select className="border border-slate-300 dark:border-slate-600 p-1 rounded-lg bg-sky-50 dark:bg-slate-900 text-[#117097] dark:text-sky-400 text-[10px] font-black outline-none" value={filaEditable.tipo_item} onChange={e => setFilaEditable({ ...filaEditable, tipo_item: e.target.value })}>
                            <option value="Consumible">🧪 Consumible</option>
                            <option value="Herramienta">🛠️ Herramienta</option>
                          </select>
                        ) : (
                          <span className={`px-2 py-0.5 rounded-md text-[9px] font-black uppercase shadow-sm border ${ins.tipo_item === 'Consumible' ? 'bg-emerald-50 dark:bg-emerald-950/80 text-emerald-700 dark:text-emerald-400 border-emerald-200 dark:border-emerald-800' : 'bg-purple-50 dark:bg-purple-950/80 text-purple-700 dark:text-purple-400 border-purple-200 dark:border-purple-800'}`}>
                            {ins.tipo_item === 'Consumible' ? '🧪 Consumible' : '🛠️ Activo'}
                          </span>
                        )}
                      </td>

                      <td className="py-2 px-3 text-center whitespace-nowrap">
                        {editandoEste ? (
                          <select className="border border-amber-300 dark:border-amber-700 bg-amber-50 dark:bg-slate-900 p-1 rounded-lg text-[10px] font-black text-amber-900 dark:text-amber-200 outline-none max-w-[150px]" value={filaEditable.ubicacion} onChange={e => setFilaEditable({ ...filaEditable, ubicacion: e.target.value })}>
                            <optgroup label="🏢 Bodegas & Áreas">
                              {listaUbicaciones.map(u => <option key={u.id} value={u.nombre}>{u.nombre}</option>)}
                            </optgroup>
                            <optgroup label="🌱 Invernaderos / Lotes">
                              {datosInvernaderos?.map(inv => <option key={inv.id} value={inv.nombre?.toUpperCase()}>{inv.nombre?.toUpperCase()}</option>)}
                            </optgroup>
                          </select>
                        ) : (
                          <span className="inline-block bg-amber-50 dark:bg-amber-950/50 text-amber-800 dark:text-amber-300 border border-amber-200 dark:border-amber-800 px-2 py-0.5 rounded-md text-[9px] font-black uppercase shadow-sm">
                            📍 {ins.ubicacion || 'BODEGA PRINCIPAL'}
                          </span>
                        )}
                      </td>

                      <td className={`py-2 px-3 text-right font-black text-xs ${estado.claseTexto} whitespace-nowrap`}>
                        {editandoEste ? (
                          <input type="number" step="any" className="border border-slate-300 dark:border-slate-600 p-1 rounded-lg w-20 text-right font-black bg-white dark:bg-slate-900 text-slate-800 dark:text-white" value={filaEditable.cantidad_actual} onChange={e => setFilaEditable({...filaEditable, cantidad_actual: e.target.value})} />
                        ) : (
                          ins.cantidad_actual
                        )}
                      </td>

                      <td className="py-2 px-3 text-center text-gray-500 dark:text-slate-400 text-[10px] uppercase tracking-wider whitespace-nowrap">
                        {editandoEste ? (
                          <select className="border border-slate-300 dark:border-slate-600 p-1 rounded-lg bg-white dark:bg-slate-900 text-[10px] font-bold text-slate-800 dark:text-white" value={filaEditable.unidad_medida} onChange={e => setFilaEditable({...filaEditable, unidad_medida: e.target.value})}>
                            {unidades.map(un => <option key={un} value={un}>{un}</option>)}
                          </select>
                        ) : (
                          ins.unidad_medida
                        )}
                      </td>

                      <td className="py-2 px-3 text-center whitespace-nowrap">
                        <span className={`px-2 py-0.5 rounded-md text-[9px] font-black uppercase ${estado.claseTabla}`}>
                          {estado.etiqueta}
                        </span>
                      </td>

                      {esAdmin && (
                        <td className="py-2 px-3 text-center whitespace-nowrap">
                          {editandoEste ? (
                            <div className="flex gap-1 justify-center">
                              <button onClick={() => guardarEdicionFila(ins.id)} className="bg-emerald-600 hover:bg-emerald-700 text-white px-2 py-1 rounded-lg text-[9px] font-black cursor-pointer shadow-sm transition-colors" title="Guardar Cambios">💾</button>
                              <button onClick={() => setIdEditando(null)} className="bg-slate-400 hover:bg-slate-500 text-white px-2 py-1 rounded-lg text-[9px] font-black cursor-pointer shadow-sm transition-colors" title="Cancelar">❌</button>
                            </div>
                          ) : (
                            <div className="flex gap-1 justify-center">
                              <button onClick={() => iniciarEdicion(ins)} className="p-1 bg-slate-700 dark:bg-slate-600 hover:bg-slate-900 dark:hover:bg-slate-500 text-white rounded-lg shadow-sm border border-slate-800 dark:border-slate-500 transition-colors text-[9px] cursor-pointer" title="Editar Registro">✏️</button>
                              <button onClick={() => eliminarInsumoCompleto(ins.id, ins.nombre_insumo)} className="p-1 bg-red-100 dark:bg-red-950/60 text-red-700 dark:text-red-400 rounded-lg hover:bg-red-700 hover:text-white transition-colors border border-red-200 dark:border-red-900 cursor-pointer" title="Eliminar definitivamente">🗑️</button>
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

      {/* 🧊 MODALES FLOTANTES (Mantenidos sin cambios de funcionalidad) */}
      
      {/* 1. MODAL: NUEVO ARTÍCULO */}
      {modalActivo === 'nuevo' && esAdmin && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 backdrop-blur-sm p-4 overflow-y-auto">
          <div className="bg-white dark:bg-slate-800 border border-slate-200 dark:border-slate-700 w-full max-w-2xl rounded-3xl shadow-2xl overflow-hidden animate-in fade-in zoom-in-95 duration-200 my-8 text-slate-800 dark:text-slate-200">
            <div className="p-5 bg-slate-900 text-white border-b border-slate-800 flex justify-between items-center">
              <div className="flex items-center gap-3">
                <span className="text-xl">➕</span>
                <div>
                  <h3 className="text-base font-black uppercase tracking-tight">Registrar Insumo o Herramienta</h3>
                  <p className="text-[10px] text-slate-400 font-medium">Creación de un nuevo elemento en el catálogo de bodega.</p>
                </div>
              </div>
              <button onClick={() => setModalActivo(null)} className="text-slate-400 hover:text-white text-lg font-black px-3 py-1 rounded-xl bg-slate-800 cursor-pointer">✕</button>
            </div>

            <div className="p-6 space-y-4 max-h-[80vh] overflow-y-auto">
              <form onSubmit={crearInsumoBauche} className="space-y-4">
                <div>
                  <label className="text-[10px] font-black text-slate-500 dark:text-slate-400 uppercase tracking-wider">Tipo de Artículo *</label>
                  <select className="w-full bg-slate-50 dark:bg-slate-900 border border-slate-200 dark:border-slate-700 text-slate-800 dark:text-white p-2.5 rounded-xl font-black text-xs outline-none focus:border-[#117097] mt-1" value={insumoForm.tipo_item} onChange={e => setInsumoForm({ ...insumoForm, tipo_item: e.target.value })}>
                    <option value="Consumible">🧪 CONSUMIBLE (Fertilizante, Fungicida, Dosis)</option>
                    <option value="Herramienta">🛠️ HERRAMIENTA / ACTIVO (Motobomba, Manguera)</option>
                  </select>
                </div>

                <div>
                  <label className="text-[10px] font-black text-slate-500 dark:text-slate-400 uppercase tracking-wider">Nombre del Artículo *</label>
                  <input type="text" className="w-full bg-slate-50 dark:bg-slate-900 border border-slate-200 dark:border-slate-700 text-slate-800 dark:text-white p-2.5 rounded-xl font-bold uppercase text-xs outline-none focus:border-[#117097] mt-1 placeholder-slate-400" value={insumoForm.nombre_insumo} onChange={e => setInsumoForm({...insumoForm, nombre_insumo: e.target.value})} placeholder="Ej: Triple 15 / Motobomba Honda" required />
                </div>

                <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                  <div>
                    <label className="text-[10px] font-black text-slate-500 dark:text-slate-400 uppercase tracking-wider">Categoría</label>
                    <select className="w-full bg-slate-50 dark:bg-slate-900 border border-slate-200 dark:border-slate-700 text-slate-800 dark:text-white p-2.5 rounded-xl font-bold text-xs outline-none focus:border-[#117097] mt-1" value={insumoForm.categoria} onChange={e => setInsumoForm({...insumoForm, categoria: e.target.value})}>
                      {categorias.map(cat => <option key={cat} value={cat}>{cat}</option>)}
                    </select>
                  </div>
                  <div>
                    <label className="text-[10px] font-black text-slate-500 dark:text-slate-400 uppercase tracking-wider">Unidad Medida</label>
                    <select className="w-full bg-slate-50 dark:bg-slate-900 border border-slate-200 dark:border-slate-700 text-slate-800 dark:text-white p-2.5 rounded-xl font-bold text-xs outline-none focus:border-[#117097] mt-1" value={insumoForm.unidad_medida} onChange={e => setInsumoForm({...insumoForm, unidad_medida: e.target.value})}>
                      {unidades.map(un => <option key={un} value={un}>{un}</option>)}
                    </select>
                  </div>
                </div>

                <div className="bg-amber-50 dark:bg-amber-950/40 p-4 rounded-2xl border border-amber-200 dark:border-amber-800 shadow-inner">
                  <label className="text-[10px] font-black text-amber-900 dark:text-amber-400 uppercase tracking-wider block mb-1">📦 Cantidad / Existencias Iniciales *</label>
                  <input type="number" step="any" className="w-full border-2 border-amber-300 dark:border-amber-700 p-2.5 rounded-xl font-black text-lg text-amber-900 dark:text-amber-200 bg-white dark:bg-slate-900 outline-none focus:border-amber-500" value={insumoForm.cantidad_inicial} onChange={e => setInsumoForm({...insumoForm, cantidad_inicial: e.target.value})} placeholder="Ej: 10" required />
                </div>

                <div>
                  <label className="text-[10px] font-black text-slate-500 dark:text-slate-400 uppercase tracking-wider">Ubicación / Localización Física</label>
                  <select className="w-full bg-slate-50 dark:bg-slate-900 border border-slate-200 dark:border-slate-700 text-slate-800 dark:text-white p-2.5 rounded-xl font-bold text-xs outline-none focus:border-[#117097] mt-1" value={insumoForm.ubicacion} onChange={e => setInsumoForm({...insumoForm, ubicacion: e.target.value})}>
                    <optgroup label="🏢 Bodegas & Áreas">
                      {listaUbicaciones.map(u => <option key={u.id} value={u.nombre}>{u.nombre}</option>)}
                    </optgroup>
                    <optgroup label="🌱 Invernaderos / Lotes">
                      {datosInvernaderos?.map(inv => <option key={inv.id} value={inv.nombre?.toUpperCase()}>{inv.nombre?.toUpperCase()}</option>)}
                    </optgroup>
                  </select>
                </div>

                {insumoForm.tipo_item === 'Consumible' ? (
                  <div className="bg-slate-50 dark:bg-slate-900/60 p-4 rounded-xl border border-slate-200 dark:border-slate-700 space-y-3">
                    <label className="flex items-center gap-2 cursor-pointer">
                      <input type="checkbox" className="w-4 h-4 accent-[#117097]" checked={insumoForm.aplica_stock} onChange={e => setInsumoForm({ ...insumoForm, aplica_stock: e.target.checked })} />
                      <span className="text-[11px] font-black text-slate-700 dark:text-slate-300 uppercase">¿Controlar Stock Mínimo para Alertas?</span>
                    </label>

                    {insumoForm.aplica_stock && (
                      <div className="pt-1 border-t border-slate-200 dark:border-slate-700 mt-2">
                        <label className="text-[10px] font-black text-slate-500 dark:text-slate-400 uppercase tracking-wider">Stock Mínimo Deseado</label>
                        <input type="number" className="w-full bg-white dark:bg-slate-800 border border-slate-200 dark:border-slate-700 text-slate-800 dark:text-white p-2.5 rounded-xl font-bold text-xs mt-1 outline-none" value={insumoForm.stock_minimo} onChange={e => setInsumoForm({...insumoForm, stock_minimo: e.target.value})} required />
                      </div>
                    )}
                  </div>
                ) : (
                  <div>
                    <label className="text-[10px] font-black text-slate-500 dark:text-slate-400 uppercase tracking-wider">Estado Inicial de la Herramienta</label>
                    <select className="w-full bg-slate-50 dark:bg-slate-900 border border-slate-200 dark:border-slate-700 text-slate-800 dark:text-white p-2.5 rounded-xl font-bold text-xs outline-none focus:border-[#117097] mt-1" value={insumoForm.estado_herramienta} onChange={e => setInsumoForm({...insumoForm, estado_herramienta: e.target.value})}>
                      {estadosHerramientaOptions.map(est => <option key={est} value={est}>{est}</option>)}
                    </select>
                  </div>
                )}

                <div className="flex justify-end gap-3 pt-4 border-t border-slate-200 dark:border-slate-700">
                  <button type="button" onClick={() => setModalActivo(null)} className="px-5 py-2.5 bg-slate-200 dark:bg-slate-700 hover:bg-slate-300 dark:hover:bg-slate-600 text-slate-700 dark:text-slate-200 font-black text-xs uppercase tracking-wider rounded-xl transition-colors cursor-pointer">Cancelar</button>
                  <button type="submit" className="px-6 py-2.5 bg-[#117097] hover:bg-[#0a4c68] text-white font-black text-xs uppercase tracking-wider rounded-xl shadow-lg transition-colors cursor-pointer">💾 Guardar Registro</button>
                </div>
              </form>

              <div className="bg-slate-50 dark:bg-slate-900/60 p-4 rounded-xl border border-slate-200 dark:border-slate-700 mt-4">
                <p className="text-[10px] font-black text-slate-600 dark:text-slate-400 uppercase tracking-wider mb-2">📍 ¿Falta una ubicación? Añádela aquí:</p>
                <div className="flex gap-2">
                  <input type="text" className="w-full border border-slate-300 dark:border-slate-600 p-2.5 bg-white dark:bg-slate-800 rounded-xl text-xs font-bold uppercase outline-none text-slate-800 dark:text-white placeholder-slate-400" value={nuevaUbicacionTexto} onChange={e => setNuevaUbicacionTexto(e.target.value)} placeholder="Ej: Cuarto de Herramientas" />
                  <button type="button" onClick={agregarNuevaUbicacion} className="px-4 bg-emerald-600 hover:bg-emerald-700 text-white font-black rounded-xl text-xs uppercase cursor-pointer transition-colors shadow">Añadir</button>
                </div>
              </div>

            </div>
          </div>
        </div>
      )}

      {/* 2. MODAL: REGISTRAR INGRESO / COMPRA */}
      {modalActivo === 'entrada' && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 backdrop-blur-sm p-4 overflow-y-auto">
          <div className="bg-white dark:bg-slate-800 border border-slate-200 dark:border-slate-700 w-full max-w-xl rounded-3xl shadow-2xl overflow-hidden animate-in fade-in zoom-in-95 duration-200 my-8 text-slate-800 dark:text-slate-200">
            <div className="p-5 bg-emerald-800 text-white border-b border-emerald-900 flex justify-between items-center">
              <div className="flex items-center gap-3">
                <span className="text-xl">📥</span>
                <div>
                  <h3 className="text-base font-black uppercase tracking-tight">Ingreso de Insumos</h3>
                  <p className="text-[10px] text-emerald-200 font-medium">Registra compras para incrementar el stock.</p>
                </div>
              </div>
              <button onClick={() => setModalActivo(null)} className="text-emerald-200 hover:text-white text-lg font-black px-3 py-1 rounded-xl bg-emerald-900 cursor-pointer">✕</button>
            </div>

            <form onSubmit={registrarEntradaBodega} className="p-6 space-y-4 max-h-[80vh] overflow-y-auto">
              <div>
                <label className="text-[10px] font-black text-slate-500 dark:text-slate-400 uppercase tracking-wider">Seleccione Insumo Consumible *</label>
                <select className="w-full bg-slate-50 dark:bg-slate-900 border border-slate-200 dark:border-slate-700 text-slate-800 dark:text-white p-3 rounded-xl font-bold text-xs outline-none focus:border-emerald-700 mt-1" value={entradaForm.insumo_id} onChange={e => setEntradaForm({...entradaForm, insumo_id: e.target.value})} required>
                  <option value="">Seleccione insumo...</option>
                  {listaInventario.filter(i => i.tipo_item === 'Consumible').map(i => (
                    <option key={i.id} value={i.id}>{i.nombre_insumo} ({i.unidad_medida}) — Stock: {i.cantidad_actual}</option>
                  ))}
                </select>
              </div>
              
              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                <div>
                  <label className="text-[10px] font-black text-emerald-700 dark:text-emerald-400 uppercase tracking-wider">Cantidad Ingresada *</label>
                  <input type="number" step="any" className="w-full bg-emerald-50 dark:bg-emerald-950/40 border border-emerald-200 dark:border-emerald-800 p-3 rounded-xl font-black text-lg text-emerald-800 dark:text-emerald-300 outline-none focus:border-emerald-500 mt-1" value={entradaForm.cantidad_ingresada} onChange={e => setEntradaForm({...entradaForm, cantidad_ingresada: e.target.value})} required />
                </div>
                <div>
                  <label className="text-[10px] font-black text-slate-500 dark:text-slate-400 uppercase tracking-wider">P. Unitario de Compra (COP)</label>
                  <input type="number" className="w-full bg-slate-50 dark:bg-slate-900 border border-slate-200 dark:border-slate-700 text-slate-800 dark:text-white p-3 rounded-xl font-bold text-lg outline-none focus:border-emerald-700 mt-1 placeholder-slate-400" value={entradaForm.precio_unitario_compra} onChange={e => setEntradaForm({...entradaForm, precio_unitario_compra: e.target.value})} placeholder="Opcional" />
                </div>
              </div>

              <div>
                <label className="text-[10px] font-black text-slate-500 dark:text-slate-400 uppercase tracking-wider">N° Comprobante / Factura</label>
                <input type="text" className="w-full bg-slate-50 dark:bg-slate-900 border border-slate-200 dark:border-slate-700 text-slate-800 dark:text-white p-3 rounded-xl font-bold outline-none focus:border-emerald-700 mt-1 uppercase placeholder-slate-400" value={entradaForm.numero_factura_comprobante} onChange={e => setEntradaForm({...entradaForm, numero_factura_comprobante: e.target.value})} placeholder="Ej: FAC-4589" />
              </div>

              <div>
                <label className="text-[10px] font-black text-slate-500 dark:text-slate-400 uppercase tracking-wider">Observaciones / Proveedor</label>
                <textarea className="w-full bg-slate-50 dark:bg-slate-900 border border-slate-200 dark:border-slate-700 text-slate-800 dark:text-white p-3 rounded-xl font-bold h-20 outline-none focus:border-emerald-700 mt-1 uppercase resize-none placeholder-slate-400" value={entradaForm.observaciones} onChange={e => setEntradaForm({...entradaForm, observaciones: e.target.value})} placeholder="Ej: Comprado en Almacén El Ganadero" />
              </div>

              <div className="flex justify-end gap-3 pt-4 border-t border-slate-200 dark:border-slate-700">
                <button type="button" onClick={() => setModalActivo(null)} className="px-5 py-2.5 bg-slate-200 dark:bg-slate-700 hover:bg-slate-300 dark:hover:bg-slate-600 text-slate-700 dark:text-slate-200 font-black text-xs uppercase tracking-wider rounded-xl transition-colors cursor-pointer">Cancelar</button>
                <button type="submit" className="px-6 py-2.5 bg-emerald-600 hover:bg-emerald-700 text-white font-black text-xs uppercase tracking-wider rounded-xl shadow-lg transition-colors cursor-pointer">📈 Confirmar Ingreso</button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* 3. MODAL: REGISTRAR SALIDA / CONSUMO */}
      {modalActivo === 'salida' && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 backdrop-blur-sm p-4 overflow-y-auto">
          <div className="bg-white dark:bg-slate-800 border border-slate-200 dark:border-slate-700 w-full max-w-xl rounded-3xl shadow-2xl overflow-hidden animate-in fade-in zoom-in-95 duration-200 my-8 text-slate-800 dark:text-slate-200">
            <div className="p-5 bg-rose-800 text-white border-b border-rose-900 flex justify-between items-center">
              <div className="flex items-center gap-3">
                <span className="text-xl">🧪</span>
                <div>
                  <h3 className="text-base font-black uppercase tracking-tight">Aplicación / Consumo</h3>
                  <p className="text-[10px] text-rose-200 font-medium">Registra el uso de insumos para descontar de bodega.</p>
                </div>
              </div>
              <button onClick={() => setModalActivo(null)} className="text-rose-200 hover:text-white text-lg font-black px-3 py-1 rounded-xl bg-rose-900 cursor-pointer">✕</button>
            </div>

            <form onSubmit={registrarSalidaInvernadero} className="p-6 space-y-4 max-h-[80vh] overflow-y-auto">
              <div>
                <label className="text-[10px] font-black text-slate-500 dark:text-slate-400 uppercase tracking-wider">Insumo Consumible a Aplicar *</label>
                <select className="w-full bg-slate-50 dark:bg-slate-900 border border-slate-200 dark:border-slate-700 text-slate-800 dark:text-white p-3 rounded-xl font-bold text-xs outline-none focus:border-rose-700 mt-1" value={salidaForm.insumo_id} onChange={e => setSalidaForm({...salidaForm, insumo_id: e.target.value})} required>
                  <option value="">Seleccione insumo...</option>
                  {listaInventario.filter(i => i.tipo_item === 'Consumible').map(i => (
                    <option key={i.id} value={i.id}>{i.nombre_insumo} — Disponible: {i.cantidad_actual} {i.unidad_medida}(s)</option>
                  ))}
                </select>
              </div>

              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                <div>
                  <label className="text-[10px] font-black text-rose-700 dark:text-rose-400 uppercase tracking-wider">Cantidad Consumida / Dosis *</label>
                  <input type="number" step="any" className="w-full bg-rose-50 dark:bg-rose-950/40 border border-rose-200 dark:border-rose-800 p-3 rounded-xl font-black text-lg text-rose-800 dark:text-rose-300 outline-none focus:border-rose-500 mt-1" value={salidaForm.cantidad_retirada} onChange={e => setSalidaForm({...salidaForm, cantidad_retirada: e.target.value})} required />
                </div>
                <div>
                  <label className="text-[10px] font-black text-slate-500 dark:text-slate-400 uppercase tracking-wider">Cultivo Destino *</label>
                  <select className="w-full bg-slate-50 dark:bg-slate-900 border border-slate-200 dark:border-slate-700 text-slate-800 dark:text-white p-3 rounded-xl font-bold text-xs outline-none focus:border-rose-700 mt-1" value={salidaForm.invernadero_id} onChange={e => setSalidaForm({...salidaForm, invernadero_id: e.target.value})} required>
                    <option value="">Seleccione Destino...</option>
                    {datosInvernaderos?.map(inv => (
                      <option key={inv.id} value={inv.id}>{inv.nombre?.toUpperCase() || inv.nombre_invernadero?.toUpperCase()}</option>
                    ))}
                  </select>
                </div>
              </div>

              <div>
                <label className="text-[10px] font-black text-slate-500 dark:text-slate-400 uppercase tracking-wider">Responsable Aplicador (Nómina) *</label>
                <select className="w-full bg-slate-50 dark:bg-slate-900 border border-slate-200 dark:border-slate-700 text-slate-800 dark:text-white p-3 rounded-xl font-bold uppercase text-xs outline-none focus:border-rose-700 mt-1" value={salidaForm.responsable} onChange={e => setSalidaForm({ ...salidaForm, responsable: e.target.value })} required>
                  <option value="">Seleccione operario...</option>
                  {listaTrabajadores.map(trab => (
                    <option key={trab.id} value={trab.nombre_completo}>{trab.nombre_completo}</option>
                  ))}
                </select>
              </div>

              <div>
                <label className="text-[10px] font-black text-slate-500 dark:text-slate-400 uppercase tracking-wider">Nota de Aplicación / Plaga</label>
                <textarea className="w-full bg-slate-50 dark:bg-slate-900 border border-slate-200 dark:border-slate-700 text-slate-800 dark:text-white p-3 rounded-xl font-bold h-20 outline-none focus:border-rose-700 mt-1 uppercase resize-none placeholder-slate-400" value={salidaForm.nota_uso} onChange={e => setSalidaForm({...salidaForm, nota_uso: e.target.value})} placeholder="Ej: Control de roya con bomba de espalda" />
              </div>

              <div className="flex justify-end gap-3 pt-4 border-t border-slate-200 dark:border-slate-700">
                <button type="button" onClick={() => setModalActivo(null)} className="px-5 py-2.5 bg-slate-200 dark:bg-slate-700 hover:bg-slate-300 dark:hover:bg-slate-600 text-slate-700 dark:text-slate-200 font-black text-xs uppercase tracking-wider rounded-xl transition-colors cursor-pointer">Cancelar</button>
                <button type="submit" className="px-6 py-2.5 bg-rose-600 hover:bg-rose-700 text-white font-black text-xs uppercase tracking-wider rounded-xl shadow-lg transition-colors cursor-pointer">📉 Confirmar Descuento</button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* 4. MODAL: REUBICAR HERRAMIENTAS */}
      {modalActivo === 'reubicar' && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 backdrop-blur-sm p-4 overflow-y-auto">
          <div className="bg-white dark:bg-slate-800 border border-slate-200 dark:border-slate-700 w-full max-w-xl rounded-3xl shadow-2xl overflow-hidden animate-in fade-in zoom-in-95 duration-200 my-8 text-slate-800 dark:text-slate-200">
            <div className="p-5 bg-purple-800 text-white border-b border-purple-900 flex justify-between items-center">
              <div className="flex items-center gap-3">
                <span className="text-xl">🔄</span>
                <div>
                  <h3 className="text-base font-black uppercase tracking-tight">Reubicar Activo</h3>
                  <p className="text-[10px] text-purple-200 font-medium">Traslado de herramientas a nuevos bloques o áreas.</p>
                </div>
              </div>
              <button onClick={() => setModalActivo(null)} className="text-purple-200 hover:text-white text-lg font-black px-3 py-1 rounded-xl bg-purple-900 cursor-pointer">✕</button>
            </div>

            <form onSubmit={reubicarHerramienta} className="p-6 space-y-4 max-h-[80vh] overflow-y-auto">
              <div>
                <label className="text-[10px] font-black text-slate-500 dark:text-slate-400 uppercase tracking-wider">Herramienta o Equipo a Mover *</label>
                <select className="w-full bg-slate-50 dark:bg-slate-900 border border-slate-200 dark:border-slate-700 text-slate-800 dark:text-white p-3 rounded-xl font-bold text-xs outline-none focus:border-purple-700 mt-1" value={reubicacionForm.herramienta_id} onChange={e => setReubicacionForm({ ...reubicacionForm, herramienta_id: e.target.value })} required>
                  <option value="">Seleccione activo...</option>
                  {listaInventario.filter(i => i.tipo_item !== 'Consumible').map(i => (
                    <option key={i.id} value={i.id}>{i.nombre_insumo} — Ubicación Actual: {i.ubicacion || 'BODEGA PRINCIPAL'}</option>
                  ))}
                </select>
              </div>

              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                <div>
                  <label className="text-[10px] font-black text-purple-700 dark:text-purple-400 uppercase tracking-wider">Nueva Ubicación Destino *</label>
                  <select className="w-full bg-purple-50 dark:bg-slate-900 border border-purple-200 dark:border-purple-700 p-3 rounded-xl font-bold text-purple-900 dark:text-purple-200 text-xs outline-none focus:border-purple-500 mt-1" value={reubicacionForm.nueva_ubicacion} onChange={e => setReubicacionForm({ ...reubicacionForm, nueva_ubicacion: e.target.value })} required>
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
                  <label className="text-[10px] font-black text-slate-500 dark:text-slate-400 uppercase tracking-wider">Asignar a Operario (Opcional)</label>
                  <select className="w-full bg-slate-50 dark:bg-slate-900 border border-slate-200 dark:border-slate-700 text-slate-800 dark:text-white p-3 rounded-xl font-bold uppercase text-xs outline-none focus:border-purple-700 mt-1" value={reubicacionForm.responsable} onChange={e => setReubicacionForm({ ...reubicacionForm, responsable: e.target.value })}>
                    <option value="">Sin responsable fijo...</option>
                    {listaTrabajadores.map(trab => (
                      <option key={trab.id} value={trab.nombre_completo}>{trab.nombre_completo}</option>
                    ))}
                  </select>
                </div>
              </div>

              <div>
                <label className="text-[10px] font-black text-slate-500 dark:text-slate-400 uppercase tracking-wider">Observaciones del Traslado</label>
                <textarea className="w-full bg-slate-50 dark:bg-slate-900 border border-slate-200 dark:border-slate-700 text-slate-800 dark:text-white p-3 rounded-xl font-bold h-20 outline-none focus:border-purple-700 mt-1 uppercase resize-none placeholder-slate-400" value={reubicacionForm.observaciones} onChange={e => setReubicacionForm({ ...reubicacionForm, observaciones: e.target.value })} placeholder="Ej: Se entrega con tanque lleno de gasolina" />
              </div>

              <div className="flex justify-end gap-3 pt-4 border-t border-slate-200 dark:border-slate-700">
                <button type="button" onClick={() => setModalActivo(null)} className="px-5 py-2.5 bg-slate-200 dark:bg-slate-700 hover:bg-slate-300 dark:hover:bg-slate-600 text-slate-700 dark:text-slate-200 font-black text-xs uppercase tracking-wider rounded-xl transition-colors cursor-pointer">Cancelar</button>
                <button type="submit" className="px-6 py-2.5 bg-purple-700 hover:bg-purple-800 text-white font-black text-xs uppercase tracking-wider rounded-xl shadow-lg transition-colors cursor-pointer">🔄 Confirmar Traslado</button>
              </div>
            </form>
          </div>
        </div>
      )}

    </div>
  );
}