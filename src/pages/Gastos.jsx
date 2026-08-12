import { useState, useEffect } from 'react';
import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';

export default function Gastos({ 
  gastoForm, 
  setGastoForm, 
  listaInvernaderos, 
  listaProveedores, 
  mostrarAlerta, 
  cargarTodo, 
  supabase, 
  datosEgresos,
  eliminarGasto,   
  imprimirGastoPDF 
}) {
  const [modalAbierto, setModalAbierto] = useState(false);
  const [busqueda, setBusqueda] = useState('');
  const [errores, setErrores] = useState({});

  const categorias = ["Mano de obra", "Insumo Agricola", "Flete", "Mto (Mantenimiento)", "S.Publicos", "Plantas", "Plasticos", "Viaticos","Arriendos", "Quincena", "Otros"];
  const unidades = ["Canastilla", "Kilo", "Bulto", "Litro", "Jornal", "Unidad", "Hora", "Otra", "Caja", "Garrafa", "Galon"];
  const formasPago = ["Efectivo", "Bre-B (Pago Inmediato)", "Bancolombia Ahorros", "Bancolombia Corriente", "Nequi", "Daviplata", "Banco de Bogotá", "Colpatria", "Davivienda", "Otro"];
  const referencias = ["FACTURA", "TICKET / TIRILLA", "RECIBO", "COTIZACIÓN", "RECIBO DE CAJA", "CUENTA DE COBRO", "OTRO"];

  const obtenerFechaLocalHoy = () => {
    const d = new Date();
    const offset = d.getTimezoneOffset();
    const localDate = new Date(d.getTime() - (offset * 60 * 1000));
    return localDate.toISOString().split('T')[0];
  };

  const limpiarFecha = (fechaStr) => {
    if (!fechaStr) return '';
    return String(fechaStr).split('T')[0];
  };

  const formatoPesos = (valor) => new Intl.NumberFormat('es-CO', { 
    style: 'currency', 
    currency: 'COP', 
    minimumFractionDigits: 0 
  }).format(valor || 0);

  // 🧮 Cálculo automático en tiempo real de Subtotal, IVA y Monto Total
  useEffect(() => {
    const cant = parseFloat(gastoForm.cantidad) || 1;
    const precio = parseFloat(gastoForm.precio_unitario) || 0;
    const subtotal = cant * precio;
    const porcIva = parseFloat(gastoForm.porcentaje_iva) || 0;
    const valorIva = (subtotal * porcIva) / 100;
    const montoTotal = subtotal + valorIva;

    if (subtotal !== gastoForm.subtotal || valorIva !== gastoForm.valor_iva || montoTotal !== gastoForm.monto) {
      setGastoForm(prev => ({
        ...prev,
        subtotal,
        valor_iva: valorIva,
        monto: montoTotal
      }));
    }
  }, [gastoForm.cantidad, gastoForm.precio_unitario, gastoForm.porcentaje_iva, setGastoForm]);

  // 🔄 Selección de proveedor con autocompletado original de banco y cuenta
  const handleCambioProveedor = (idSeleccionado) => {
    const prov = listaProveedores?.find(p => p.id?.toString() === idSeleccionado?.toString());
    
    setGastoForm(prev => ({
      ...prev,
      proveedor_id: idSeleccionado,
      nombre_proveedor: prov ? (prov.nombre || prov.nombre_completo || '') : '',
      nit_cc: prov ? (prov.nit || prov.nit_cc || '') : '',
      direccion: prov ? (prov.dir || prov.direccion || '') : '',
      telefono: prov ? (prov.tel || prov.telefono || '') : '',
      forma_pago: (!prev.id_editando && prov && prov.banco) ? prov.banco : prev.forma_pago,
      numero_cuenta: prov ? (prov.numero_cuenta || '') : prev.numero_cuenta
    }));
    setErrores(prev => ({ ...prev, proveedor: null }));
  };

  const abrirModalNuevo = () => {
    setGastoForm({ 
      id_editando: null,
      referencia: 'FACTURA',
      invernadero_id: '', 
      descripcion: '', 
      monto: 0, 
      subtotal: 0,
      porcentaje_iva: 0,
      valor_iva: 0,
      categoria: 'Insumo Agricola', 
      proveedor_id: '', 
      nombre_proveedor: '',
      nit_cc: '',
      direccion: '',
      telefono: '',
      numero_comprobante: '', 
      nota: '', 
      fecha: obtenerFechaLocalHoy(), 
      cantidad: 1, 
      unidad_medida: 'Unidad', 
      precio_unitario: '',
      forma_pago: 'Efectivo',
      numero_cuenta: ''
    });
    setErrores({});
    setModalAbierto(true);
  };

  // 🛠️ FUNCIÓN DE EDICIÓN MEJORADA PARA SOPORTAR PROVEEDORES LIBRES Y DE TELEGRAM
  const prepararEdicionGasto = (g) => {
    setGastoForm({ 
      id_editando: g.id,
      referencia: g.referencia || 'FACTURA',
      invernadero_id: g.invernadero_id || '', 
      descripcion: g.descripcion || '', 
      monto: g.monto || 0, 
      subtotal: g.subtotal || g.monto || 0,
      porcentaje_iva: g.porcentaje_iva || 0,
      valor_iva: g.valor_iva || 0,
      categoria: g.categoria || 'Insumo Agricola', 
      proveedor_id: g.proveedor_id || '', 
      nombre_proveedor: g.nombre_proveedor || '',
      nit_cc: g.nit_cc || '',
      numero_comprobante: g.numero_comprobante || '', 
      nota: g.nota || '', 
      fecha: limpiarFecha(g.fecha || g.fecha_gasto) || obtenerFechaLocalHoy(), 
      cantidad: g.cantidad || 1, 
      unidad_medida: g.unidad_medida || 'Unidad', 
      precio_unitario: g.precio_unitario || '',
      forma_pago: g.forma_pago || 'Efectivo',
      numero_cuenta: g.numero_cuenta || ''
    });
    setErrores({});
    setModalAbierto(true);
  };

  // --- 🛑 VALIDACIÓN Y ENVÍO ---
  const validarYGuardarGasto = async (e) => {
    if (e && e.preventDefault) e.preventDefault();

    const nuevosErrores = {};
    const fechaLimpia = (gastoForm.fecha || '').trim();
    const comprobanteLimpio = (gastoForm.numero_comprobante || '').trim();
    const invernaderoLimpio = (gastoForm.invernadero_id || '').toString().trim();
    const descripcionLimpia = (gastoForm.descripcion || '').trim();
    const cantidadNum = parseFloat(gastoForm.cantidad) || 0;
    const precioNum = parseFloat(gastoForm.precio_unitario) || 0;

    if (!fechaLimpia) nuevosErrores.fecha = "Seleccione una fecha de pago.";
    if (!comprobanteLimpio) nuevosErrores.comprobante = "Ingrese el N° de factura.";
    if (!invernaderoLimpio) nuevosErrores.invernadero = "Seleccione un invernadero.";
    if (!descripcionLimpia) nuevosErrores.descripcion = "Ingrese el concepto del gasto.";
    if (cantidadNum <= 0) nuevosErrores.cantidad = "Introduzca una cantidad válida.";
    if (precioNum <= 0) nuevosErrores.precio = "Introduzca un precio válido.";

    if (Object.keys(nuevosErrores).length > 0) {
      setErrores(nuevosErrores);
      return;
    }

    setErrores({});

    const payload = {
      referencia: gastoForm.referencia || 'FACTURA',
      descripcion: descripcionLimpia.toUpperCase(),
      categoria: gastoForm.categoria,
      monto: gastoForm.monto || (cantidadNum * precioNum),
      subtotal: gastoForm.subtotal || (cantidadNum * precioNum),
      porcentaje_iva: parseFloat(gastoForm.porcentaje_iva) || 0,
      valor_iva: gastoForm.valor_iva || 0,
      invernadero_id: invernaderoLimpio,
      proveedor_id: gastoForm.proveedor_id || null,
      nombre_proveedor: gastoForm.nombre_proveedor || '',
      nit_cc: gastoForm.nit_cc || '',
      direccion: gastoForm.direccion || '',
      telefono: gastoForm.telefono || '',
      numero_comprobante: comprobanteLimpio.toUpperCase(),
      nota: gastoForm.nota ? gastoForm.nota.trim() : null,
      fecha: fechaLimpia,
      cantidad: cantidadNum,
      unidad_medida: gastoForm.unidad_medida || 'Unidad',
      precio_unitario: precioNum,
      forma_pago: gastoForm.forma_pago || 'Efectivo',
      numero_cuenta: gastoForm.numero_cuenta || null
    };

    try {
      if (gastoForm.id_editando) {
        const { error } = await supabase.from('egresos').update(payload).eq('id', gastoForm.id_editando);
        if (error) throw error;
        mostrarAlerta("Gasto actualizado correctamente", "exito");
      } else {
        const { error } = await supabase.from('egresos').insert([payload]);
        if (error) throw error;
        mostrarAlerta("Gasto registrado correctamente", "exito");
      }

      setModalAbierto(false);
      cargarTodo();
    } catch (error) {
      console.error("Error al guardar gasto:", error);
      mostrarAlerta("Error al guardar gasto: " + error.message, "error");
    }
  };

  // --- 📊 EXPORTAR A EXCEL ACTUALIZADO CON TIPO / REFERENCIA ---
  const exportarAExcel = async () => {
    if (!datosEgresos || datosEgresos.length === 0) {
      mostrarAlerta("No hay datos de gastos para exportar", "error");
      return;
    }

    try {
      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet('Control de Gastos');

      worksheet.columns = [
        { header: 'FECHA GASTO', key: 'fecha', width: 15 },
        { header: 'TIPO DOC', key: 'referencia', width: 18 },
        { header: 'COMPROBANTE N°', key: 'comprobante', width: 18 },
        { header: 'INVERNADERO', key: 'invernadero', width: 18 },
        { header: 'PROVEEDOR', key: 'proveedor', width: 25 },
        { header: 'NIT / CC', key: 'nit', width: 16 },
        { header: 'CATEGORÍA', key: 'categoria', width: 20 },
        { header: 'FORMA DE PAGO', key: 'forma_pago', width: 22 },
        { header: 'N° CUENTA / CELULAR', key: 'cuenta', width: 20 },
        { header: 'DESCRIPCIÓN / DETALLE', key: 'descripcion', width: 35 },
        { header: 'CANTIDAD', key: 'cantidad', width: 12 },
        { header: 'UNIDAD MEDIDA', key: 'unidad', width: 16 },
        { header: 'PRECIO UNITARIO', key: 'precio', width: 18 },
        { header: 'MONTO TOTAL', key: 'monto', width: 18 },
        { header: 'NOTA / OBSERVACIONES', key: 'nota', width: 30 }
      ];

      datosEgresos.forEach((g) => {
        const proveedor = g.nombre_proveedor || g.proveedores?.nombre_completo || g.proveedores?.nombre || 'Particular';
        const nit = g.nit_cc || g.proveedores?.nit_cc || 'N/A';
        const invernadero = g.nombre_invernadero || g.invernaderos?.nombre || 'General';

        worksheet.addRow({
          fecha: limpiarFecha(g.fecha || g.fecha_gasto),
          referencia: g.referencia || 'FACTURA',
          comprobante: g.numero_comprobante || 'S/N',
          invernadero: invernadero.toUpperCase(),
          proveedor: proveedor.toUpperCase(),
          nit: nit,
          categoria: (g.categoria || 'Sin Categoría').toUpperCase(),
          forma_pago: (g.forma_pago || 'Efectivo').toUpperCase(),
          cuenta: g.numero_cuenta || 'N/A',
          descripcion: g.descripcion || '',
          cantidad: parseFloat(g.cantidad) || 0,
          unidad: g.unidad_medida || 'Unidad',
          precio: parseFloat(g.precio_unitario) || 0,
          monto: parseFloat(g.monto) || 0,
          nota: g.nota || ''
        });
      });

      const totalRowNumber = worksheet.rowCount + 1;
      const ultimaFilaDatos = worksheet.rowCount;

      const totalRow = worksheet.addRow({
        descripcion: 'TOTAL GENERAL DE GASTOS:',
        monto: { formula: `=SUM(N2:N${ultimaFilaDatos})` } 
      });

      const headerRow = worksheet.getRow(1);
      headerRow.height = 24;
      headerRow.eachCell((cell) => {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF117097' } }; 
        cell.font = { name: 'Arial', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
      });

      worksheet.eachRow((row, rowNumber) => {
        if (rowNumber === 1 || rowNumber === totalRowNumber) return; 
        row.height = 20;
        const esCebra = rowNumber % 2 === 0;
        row.eachCell((cell, colNumber) => {
          cell.font = { name: 'Arial', size: 9 };
          if (esCebra) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEBF5FB' } }; 
          
          if ([1, 2, 3, 6, 8, 9].includes(colNumber)) cell.alignment = { vertical: 'middle', horizontal: 'center' };
          else if ([11, 13, 14].includes(colNumber)) cell.alignment = { vertical: 'middle', horizontal: 'right' };
          else cell.alignment = { vertical: 'middle', horizontal: 'left' };
          
          if (colNumber === 11) cell.numFmt = '#,##0';
          if (colNumber === 13 || colNumber === 14) cell.numFmt = '"$"#,##0';
        });
      });

      totalRow.height = 22;
      totalRow.eachCell((cell, colNumber) => {
        cell.font = { name: 'Arial', size: 10, bold: true, color: { argb: 'FF0A4C68' } }; 
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD6EEFC' } }; 
        
        cell.border = {
          top: { style: 'thin', color: { argb: 'FF117097' } },
          bottom: { style: 'double', color: { argb: 'FF117097' } }
        };

        if (colNumber === 10) cell.alignment = { vertical: 'middle', horizontal: 'right' }; 
        if (colNumber === 14) {
          cell.alignment = { vertical: 'middle', horizontal: 'right' }; 
          cell.numFmt = '"$"#,##0'; 
        }
      });

      worksheet.autoFilter = { from: { row: 1, column: 1 }, to: { row: ultimaFilaDatos, column: worksheet.columnCount } };
      
      const buffer = await workbook.xlsx.writeBuffer();
      const fechaHoy = obtenerFechaLocalHoy();
      const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      saveAs(blob, `BITACORA_GASTOS_${fechaHoy}.xlsx`);
      
      if (typeof mostrarAlerta === "function") mostrarAlerta("Reporte de gastos generado con éxito", "exito");
    } catch (error) {
      console.error("Error al exportar Excel:", error);
    }
  };

  // 🔍 Filtrado de datos en tiempo real para la tabla moderna
  const datosFiltrados = datosEgresos?.filter(g => {
    const texto = busqueda.toLowerCase();
    const prov = (g.nombre_proveedor || g.proveedores?.nombre_completo || '').toLowerCase();
    const doc = (g.numero_comprobante || '').toLowerCase();
    const desc = (g.descripcion || '').toLowerCase();
    const cat = (g.categoria || '').toLowerCase();
    const ref = (g.referencia || '').toLowerCase();
    return prov.includes(texto) || doc.includes(texto) || desc.includes(texto) || cat.includes(texto) || ref.includes(texto);
  }) || [];

  const totalFiltrado = datosFiltrados.reduce((acc, g) => acc + (parseFloat(g.monto) || 0), 0);

  return (
    <div className="space-y-6 pb-20 text-slate-800 dark:text-slate-200 font-sans transition-colors duration-300">
      
      {/* 🚀 CABECERA PRINCIPAL MODERNA */}
      <div className="bg-white dark:bg-slate-800 p-6 rounded-3xl shadow-xl border border-slate-200 dark:border-slate-700 flex flex-col md:flex-row justify-between items-center gap-4 transition-colors duration-300">
        <div className="space-y-1">
          <div className="flex items-center gap-3">
            <span className="text-2xl p-2 bg-[#117097]/10 dark:bg-sky-500/20 rounded-xl text-[#117097] dark:text-sky-400">📑</span>
            <h2 className="text-xl font-black tracking-tight text-slate-900 dark:text-white uppercase">Control de Gastos y Facturación</h2>
          </div>
          <p className="text-xs text-slate-500 dark:text-slate-400 font-medium">Registro de egresos y facturas electrónicas de proveedores en INTEPE.</p>
        </div>
        
        <div className="flex items-center gap-3 w-full md:w-auto">
          <button onClick={exportarAExcel} className="flex-1 md:flex-none px-4 py-2.5 bg-emerald-600 hover:bg-emerald-700 text-white font-black text-xs uppercase tracking-wider rounded-xl shadow-md transition-all flex items-center justify-center gap-2 cursor-pointer">
            <span>📊</span> Exportar Excel
          </button>
          <button onClick={abrirModalNuevo} className="flex-1 md:flex-none px-5 py-2.5 bg-rose-600 hover:bg-rose-700 text-white font-black text-xs uppercase tracking-wider rounded-xl shadow-md transition-all flex items-center justify-center gap-2 cursor-pointer">
            <span>+</span> Registrar Gasto
          </button>
        </div>
      </div>

      {/* 🔍 BARRA DE BÚSQUEDA Y TOTALES */}
      <div className="bg-white dark:bg-slate-800 p-4 rounded-2xl shadow-md border border-slate-200 dark:border-slate-700 flex flex-col md:flex-row justify-between items-center gap-4 transition-colors duration-300">
        <div className="relative w-full md:w-96">
          <span className="absolute inset-y-0 left-0 flex items-center pl-3 text-slate-400">🔍</span>
          <input 
            type="text" 
            placeholder="Buscar por factura, proveedor, concepto..." 
            className="w-full bg-slate-50 dark:bg-slate-700/50 border border-slate-200 dark:border-slate-600 text-xs text-slate-800 dark:text-white rounded-xl pl-10 pr-4 py-2.5 outline-none focus:border-[#117097] dark:focus:border-sky-500 font-bold placeholder-slate-400"
            value={busqueda}
            onChange={e => setBusqueda(e.target.value)}
          />
        </div>
        <div className="text-right w-full md:w-auto bg-sky-50 dark:bg-sky-950/40 px-5 py-2 rounded-xl border border-sky-100 dark:border-sky-900">
          <p className="text-[9px] font-black text-slate-400 uppercase tracking-widest">Total Filtrado</p>
          <p className="text-base font-black text-[#117097] dark:text-sky-400 tracking-tight">{formatoPesos(totalFiltrado)}</p>
        </div>
      </div>

      {/* 📊 TABLA DE GASTOS ESTILO MODERNO */}
      <div className="bg-white dark:bg-slate-800 rounded-3xl shadow-xl overflow-hidden border border-slate-200 dark:border-slate-700 transition-colors duration-300">
        <div className="overflow-x-auto max-h-[600px]">
          <table className="w-full text-left text-xs border-collapse">
            <thead>
              <tr className="bg-slate-100 dark:bg-slate-700/80 text-slate-700 dark:text-slate-300 uppercase font-black text-[10px] tracking-wider border-b border-slate-200 dark:border-slate-700 sticky top-0">
                <th className="p-4">Fecha / Factura</th>
                <th className="p-4 text-center">Invernadero</th>
                <th className="p-4">Concepto / Categoría</th>
                <th className="p-4 text-center">Ref / Tipo</th>
                <th className="p-4">Proveedor / Beneficiario</th>
                <th className="p-4 text-center">Forma de Pago</th>
                <th className="p-4 text-right">Monto Total</th>
                <th className="p-4 text-center">Acciones</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100 dark:divide-slate-700/60 font-bold text-slate-700 dark:text-slate-300">
              {datosFiltrados.length === 0 ? (
                <tr>
                  <td colSpan="8" className="p-8 text-center text-slate-400 dark:text-slate-500 font-bold italic">No se encontraron registros de gastos.</td>
                </tr>
              ) : (
                datosFiltrados.map((g, index) => {
                  const nombreProveedor = g.nombre_proveedor || g.proveedores?.nombre_completo || g.proveedores?.nombre || 'Particular / Otros';
                  return (
                    <tr key={g.id} className={`${index % 2 === 0 ? 'bg-white dark:bg-slate-800' : 'bg-slate-50/50 dark:bg-slate-800/60'} hover:bg-sky-50/50 dark:hover:bg-slate-700/50 transition-colors border-l-8 border-[#117097] dark:border-sky-600`}>
                      <td className="p-4 whitespace-nowrap">
                        <div className="font-black text-slate-900 dark:text-white">{limpiarFecha(g.fecha || g.fecha_gasto)}</div>
                        <div className="text-[10px] text-[#117097] dark:text-sky-400 font-black mt-0.5">{g.numero_comprobante ? `DOC: ${g.numero_comprobante.toUpperCase()}` : 'S/N'}</div>
                      </td>
                      <td className="p-4 text-center whitespace-nowrap">
                        <span className="bg-slate-700 text-white px-2 py-0.5 rounded text-[9px] font-black uppercase shadow-sm">
                          {g.invernaderos?.nombre || 'Gral'}
                        </span>
                      </td>
                     <td className="p-4 font-bold text-slate-800 dark:text-slate-200 max-w-[300px] whitespace-normal break-words">
                        <p className="uppercase font-black text-slate-900 dark:text-white leading-snug">{g.descripcion}</p>
                        <p className="text-[9px] text-[#117097] dark:text-sky-400 font-black uppercase italic mt-1">📌 {g.categoria || 'Varios'}</p>
                      </td>
                      <td className="p-4 text-center whitespace-nowrap">
                        <span className="bg-indigo-50 dark:bg-indigo-950/60 text-indigo-700 dark:text-indigo-400 border border-indigo-200 dark:border-indigo-800 px-2 py-1 rounded-md text-[9px] font-black uppercase shadow-sm">
                          {g.referencia || 'FACTURA'}
                        </span>
                      </td>
                      <td className="p-4 font-bold text-slate-700 dark:text-slate-300">
                        <p className="uppercase text-xs font-black text-slate-800 dark:text-white">{nombreProveedor}</p>
                        <p className="text-[10px] text-slate-400 dark:text-slate-500 font-bold">NIT: {g.nit_cc || 'S/N'}</p>
                      </td>
                      <td className="p-4 text-center whitespace-nowrap">
                        <span className={`inline-block px-2 py-0.5 rounded text-[8px] uppercase tracking-wider font-black shadow-sm ${
                          g.forma_pago === 'Efectivo' 
                            ? 'bg-emerald-100 dark:bg-emerald-950/80 text-emerald-800 dark:text-emerald-400' 
                            : g.forma_pago?.toLowerCase().includes('bre-b') 
                            ? 'bg-cyan-100 dark:bg-cyan-950/80 text-cyan-800 dark:text-cyan-400 border border-cyan-200 dark:border-cyan-800' 
                            : g.forma_pago?.toLowerCase().includes('nequi') 
                            ? 'bg-purple-100 dark:bg-purple-950/80 text-purple-800 dark:text-purple-400' 
                            : g.forma_pago?.toLowerCase().includes('daviplata')
                            ? 'bg-red-100 dark:bg-red-950/80 text-red-800 dark:text-red-400'
                            : 'bg-blue-100 dark:bg-blue-950/80 text-blue-800 dark:text-blue-400'
                        }`}>
                          {g.forma_pago?.toLowerCase().includes('bre-b') ? '⚡ Bre-B' : `${g.forma_pago || 'Efectivo'}`}
                        </span>
                        {g.forma_pago !== 'Efectivo' && g.numero_cuenta && (
                          <p className="text-[9px] text-slate-500 dark:text-slate-400 font-black tracking-tight mt-1 bg-slate-100 dark:bg-slate-700 border border-slate-200 dark:border-slate-600 rounded px-1.5 py-0.5 max-w-[120px] mx-auto truncate" title={g.numero_cuenta}>
                            #{g.numero_cuenta}
                          </p>
                        )}
                      </td>
                      <td className="p-4 text-right font-black text-[#117097] dark:text-sky-400 text-sm whitespace-nowrap">
                        {formatoPesos(g.monto)}
                      </td>
                      <td className="p-4 whitespace-nowrap">
                        <div className="flex gap-1.5 justify-center">
                          <button onClick={() => prepararEdicionGasto(g)} className="px-2.5 py-1 bg-slate-700 dark:bg-slate-600 text-white rounded-lg hover:bg-slate-900 dark:hover:bg-slate-500 transition-colors flex items-center gap-1 border border-slate-800 dark:border-slate-500 shadow-md text-[9px] font-black cursor-pointer" title="Editar">
                            <span>✏️</span><span>EDITAR</span>
                          </button>
                          <button onClick={() => imprimirGastoPDF(g)} className="px-2.5 py-1 bg-slate-800 dark:bg-slate-700 text-white rounded-lg hover:bg-black dark:hover:bg-slate-600 transition-colors flex items-center gap-1 border border-slate-900 dark:border-slate-600 shadow-md text-[9px] font-black cursor-pointer" title="Imprimir PDF">
                            <span>🖨️</span><span>PDF</span>
                          </button>
                          <button onClick={() => eliminarGasto(g.id)} className="p-1.5 bg-red-100 dark:bg-red-950/60 text-red-700 dark:text-red-400 rounded-lg hover:bg-red-700 hover:text-white transition-colors border border-red-200 dark:border-red-900 cursor-pointer" title="Eliminar">
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

      {/* 🧊 MODAL FLOTANTE PARA NUEVO / EDITAR GASTO */}
      {modalAbierto && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 backdrop-blur-sm p-4 overflow-y-auto">
          <div className="bg-white dark:bg-slate-800 border border-slate-200 dark:border-slate-700 w-full max-w-2xl rounded-3xl shadow-2xl overflow-hidden animate-in fade-in zoom-in-95 duration-200 my-8 text-slate-800 dark:text-slate-200">
            
            <div className="p-5 bg-slate-900 text-white border-b border-slate-800 flex justify-between items-center">
              <div className="flex items-center gap-3">
                <span className="text-xl">📑</span>
                <div>
                  <h3 className="text-base font-black uppercase tracking-tight">
                    {gastoForm.id_editando ? "Editar Registro de Gasto" : "Nuevo Registro de Gasto"}
                  </h3>
                  <p className="text-[10px] text-slate-400 font-medium">Información contable y de proveedores en INTEPE.</p>
                </div>
              </div>
              <button onClick={() => setModalAbierto(false)} className="text-slate-400 hover:text-white text-lg font-black px-3 py-1 rounded-xl bg-slate-800 cursor-pointer">✕</button>
            </div>

            <form onSubmit={validarYGuardarGasto} className="p-6 space-y-4 max-h-[80vh] overflow-y-auto">
              
              <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                <div>
                  <label className="text-[10px] font-black text-slate-500 dark:text-slate-400 uppercase tracking-wider">Referencia (Tipo) *</label>
                  <select className="w-full bg-slate-50 dark:bg-slate-900 border border-slate-200 dark:border-slate-700 text-xs text-slate-800 dark:text-white p-2.5 rounded-xl outline-none focus:border-[#117097] font-bold mt-1"
                    value={gastoForm.referencia || 'FACTURA'} onChange={e => setGastoForm({...gastoForm, referencia: e.target.value})}>
                    {referencias.map(r => <option key={r} value={r}>{r}</option>)}
                  </select>
                </div>
                <div>
                  <label className="text-[10px] font-black text-slate-500 dark:text-slate-400 uppercase tracking-wider">N° Doc / Factura *</label>
                  <input type="text" placeholder="Ej. FE-1042" 
                    className={`w-full bg-slate-50 dark:bg-slate-900 border ${errores.comprobante ? 'border-red-500 bg-red-50 dark:bg-red-950/50' : 'border-slate-200 dark:border-slate-700'} text-xs text-slate-800 dark:text-white p-2.5 rounded-xl outline-none focus:border-[#117097] font-bold mt-1 uppercase placeholder-slate-400`}
                    value={gastoForm.numero_comprobante || ''} onChange={e => { setGastoForm({...gastoForm, numero_comprobante: e.target.value}); setErrores({...errores, comprobante: null}); }} />
                  {errores.comprobante && <p className="text-[9px] text-red-500 font-bold mt-1">{errores.comprobante}</p>}
                </div>
                <div>
                  <label className="text-[10px] font-black text-slate-500 dark:text-slate-400 uppercase tracking-wider">Fecha de Emisión *</label>
                  <input type="date" 
                    className="w-full bg-slate-50 dark:bg-slate-900 border border-slate-200 dark:border-slate-700 text-xs text-slate-800 dark:text-white p-2.5 rounded-xl outline-none focus:border-[#117097] font-bold mt-1"
                    value={gastoForm.fecha || obtenerFechaLocalHoy()} onChange={e => setGastoForm({...gastoForm, fecha: e.target.value})} />
                </div>
              </div>

              {/* 👥 SELECCIÓN DE PROVEEDOR ORIGINAL CON AUTOCOMPLETADO */}
              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                <div>
                  <label className="text-[10px] font-black text-slate-500 dark:text-slate-400 uppercase tracking-wider">Proveedor / Beneficiario *</label>
                  <select className="w-full bg-slate-50 dark:bg-slate-900 border border-slate-200 dark:border-slate-700 text-xs text-slate-800 dark:text-white p-2.5 rounded-xl outline-none focus:border-[#117097] font-bold mt-1"
                    value={gastoForm.proveedor_id || ''} onChange={e => handleCambioProveedor(e.target.value)}>
                    <option value="">{gastoForm.nombre_proveedor ? `Registrado: ${gastoForm.nombre_proveedor}` : 'Particular / Otros'}</option>
                    {listaProveedores?.map(p => <option key={p.id} value={p.id}>{p.nombre || p.nombre_completo}</option>)}
                  </select>
                </div>
                <div>
                  <label className="text-[10px] font-black text-slate-500 dark:text-slate-400 uppercase tracking-wider">NIT / CC</label>
                  <input type="text" placeholder="NIT o Cédula" 
                    className="w-full bg-slate-50 dark:bg-slate-900 border border-slate-200 dark:border-slate-700 text-xs text-slate-800 dark:text-white p-2.5 rounded-xl outline-none focus:border-[#117097] font-bold mt-1 placeholder-slate-400"
                    value={gastoForm.nit_cc || ''} onChange={e => setGastoForm({...gastoForm, nit_cc: e.target.value})} />
                </div>
              </div>

              {/* 💳 FORMA DE PAGO Y CUENTA ORIGINALES */}
              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                <div>
                  <label className="text-[10px] font-black text-slate-500 dark:text-slate-400 uppercase tracking-wider">Forma de Pago *</label>
                  <select className="w-full bg-slate-50 dark:bg-slate-900 border border-slate-200 dark:border-slate-700 text-xs text-slate-800 dark:text-white p-2.5 rounded-xl outline-none focus:border-[#117097] font-bold mt-1 text-[#117097] dark:text-sky-400" 
                    value={gastoForm.forma_pago || 'Efectivo'} onChange={e => setGastoForm({...gastoForm, forma_pago: e.target.value})}>
                    {formasPago.map(fp => <option key={fp} value={fp}>{fp}</option>)}
                  </select>
                </div>
                <div>
                  <label className="text-[10px] font-black text-slate-500 dark:text-slate-400 uppercase tracking-wider">N° Cuenta / Celular</label>
                  <input type="text" placeholder="N° Cuenta / Celular" 
                    className="w-full bg-amber-50/40 dark:bg-amber-950/40 border border-dashed border-amber-300 dark:border-amber-700 text-xs text-[#117097] dark:text-sky-400 p-2.5 rounded-xl outline-none font-black mt-1 text-center placeholder-slate-400"
                    value={gastoForm.numero_cuenta || ''} onChange={e => setGastoForm({...gastoForm, numero_cuenta: e.target.value})} />
                </div>
              </div>

              <div>
                <label className="text-[10px] font-black text-slate-500 dark:text-slate-400 uppercase tracking-wider">Concepto / Detalle del Gasto *</label>
                <input type="text" placeholder="Ej: Compra de abono o insumos..." 
                  className={`w-full bg-slate-50 dark:bg-slate-900 border ${errores.descripcion ? 'border-red-500 bg-red-50 dark:bg-red-950/50' : 'border-slate-200 dark:border-slate-700'} text-xs text-slate-800 dark:text-white p-2.5 rounded-xl outline-none focus:border-[#117097] font-bold mt-1 placeholder-slate-400`}
                  value={gastoForm.descripcion || ''} onChange={e => { setGastoForm({...gastoForm, descripcion: e.target.value}); setErrores({...errores, descripcion: null}); }} />
                {errores.descripcion && <p className="text-[9px] text-red-500 font-bold mt-1">{errores.descripcion}</p>}
              </div>

              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                <div>
                  <label className="text-[10px] font-black text-slate-500 dark:text-slate-400 uppercase tracking-wider">Invernadero Asignado *</label>
                  <select className={`w-full bg-slate-50 dark:bg-slate-900 border ${errores.invernadero ? 'border-red-500 bg-red-50 dark:bg-red-950/50' : 'border-slate-200 dark:border-slate-700'} text-xs text-slate-800 dark:text-white p-2.5 rounded-xl outline-none focus:border-[#117097] font-bold mt-1`}
                    value={gastoForm.invernadero_id || ''} onChange={e => { setGastoForm({...gastoForm, invernadero_id: e.target.value}); setErrores({...errores, invernadero: null}); }}>
                    <option value="">Seleccionar bloque...</option>
                    {listaInvernaderos?.filter(i => i.activo !== false).map(i => <option key={i.id} value={i.id}>{i.nombre}</option>)}
                  </select>
                  {errores.invernadero && <p className="text-[9px] text-red-500 font-bold mt-1">{errores.invernadero}</p>}
                </div>
                <div>
                  <label className="text-[10px] font-black text-slate-500 dark:text-slate-400 uppercase tracking-wider">Categoría *</label>
                  <select className="w-full bg-slate-50 dark:bg-slate-900 border border-slate-200 dark:border-slate-700 text-xs text-slate-800 dark:text-white p-2.5 rounded-xl outline-none focus:border-[#117097] font-bold mt-1"
                    value={gastoForm.categoria || 'Insumo Agricola'} onChange={e => setGastoForm({...gastoForm, categoria: e.target.value})}>
                    {categorias.map(c => <option key={c} value={c}>{c}</option>)}
                  </select>
                </div>
              </div>

              {/* 🧮 BLOQUE DE MONTOS, CANTIDAD E IVA */}
              <div className="bg-slate-50 dark:bg-slate-900/60 p-4 rounded-2xl border border-slate-200 dark:border-slate-700 space-y-3 shadow-inner">
                <div className="grid grid-cols-1 md:grid-cols-3 gap-3">
                  <div>
                    <label className="text-[9px] font-black text-slate-500 dark:text-slate-400 uppercase">Cantidad *</label>
                    <input type="number" min="1" 
                      className={`w-full bg-white dark:bg-slate-800 border ${errores.cantidad ? 'border-red-500 bg-red-50 dark:bg-red-950/50' : 'border-slate-200 dark:border-slate-700'} text-xs text-slate-800 dark:text-white p-2 rounded-xl font-black mt-1 outline-none`}
                      value={gastoForm.cantidad || 1} onChange={e => { setGastoForm({...gastoForm, cantidad: e.target.value}); setErrores({...errores, cantidad: null}); }} />
                    {errores.cantidad && <p className="text-[9px] text-red-500 font-bold mt-1">{errores.cantidad}</p>}
                  </div>
                  <div>
                    <label className="text-[9px] font-black text-slate-500 dark:text-slate-400 uppercase">Unidad de Medida</label>
                    <select className="w-full bg-white dark:bg-slate-800 border border-slate-200 dark:border-slate-700 text-xs text-slate-800 dark:text-white p-2 rounded-xl font-bold mt-1 outline-none"
                      value={gastoForm.unidad_medida || 'Unidad'} onChange={e => setGastoForm({...gastoForm, unidad_medida: e.target.value})}>
                      {unidades.map(u => <option key={u} value={u}>{u}</option>)}
                    </select>
                  </div>
                  <div>
                    <label className="text-[9px] font-black text-slate-500 dark:text-slate-400 uppercase">Valor Unitario ($) *</label>
                    <input type="number" placeholder="0" 
                      className={`w-full bg-white dark:bg-slate-800 border ${errores.precio ? 'border-red-500 bg-red-50 dark:bg-red-950/50' : 'border-slate-200 dark:border-slate-700'} text-xs text-[#117097] dark:text-sky-400 p-2 rounded-xl font-black mt-1 outline-none placeholder-slate-400`}
                      value={gastoForm.precio_unitario || ''} onChange={e => { setGastoForm({...gastoForm, precio_unitario: e.target.value}); setErrores({...errores, precio: null}); }} />
                    {errores.precio && <p className="text-[9px] text-red-500 font-bold mt-1">{errores.precio}</p>}
                  </div>
                </div>

                <div className="grid grid-cols-1 md:grid-cols-3 gap-3 pt-2 border-t border-slate-200 dark:border-slate-700">
                  <div>
                    <label className="text-[9px] font-black text-slate-500 dark:text-slate-400 uppercase">% IVA</label>
                    <select className="w-full bg-white dark:bg-slate-800 border border-slate-200 dark:border-slate-700 text-xs text-slate-800 dark:text-white p-2 rounded-xl font-bold mt-1 outline-none"
                      value={gastoForm.porcentaje_iva || 0} onChange={e => setGastoForm({...gastoForm, porcentaje_iva: e.target.value})}>
                      <option value="0">0%</option>
                      <option value="5">5%</option>
                      <option value="19">19%</option>
                    </select>
                  </div>
                  <div>
                    <label className="text-[9px] font-black text-slate-500 dark:text-slate-400 uppercase">Valor IVA ($)</label>
                    <input type="text" readOnly 
                      className="w-full bg-slate-100 dark:bg-slate-900 border border-slate-200 dark:border-slate-700 text-xs text-slate-600 dark:text-slate-400 p-2 rounded-xl font-black mt-1 outline-none cursor-not-allowed"
                      value={formatoPesos(gastoForm.valor_iva)} />
                  </div>
                  <div>
                    <label className="text-[9px] font-black text-[#117097] dark:text-sky-400 uppercase">Monto Total ($)</label>
                    <input type="text" readOnly 
                      className="w-full bg-sky-50 dark:bg-sky-950/50 border border-sky-200 dark:border-sky-800 text-xs text-[#117097] dark:text-sky-400 p-2 rounded-xl font-black mt-1 outline-none cursor-not-allowed"
                      value={formatoPesos(gastoForm.monto)} />
                  </div>
                </div>
              </div>

              <div>
                <label className="text-[10px] font-black text-slate-500 dark:text-slate-400 uppercase tracking-wider">Notas Adicionales</label>
                <textarea rows="2" placeholder="Observaciones o centro de costos..." 
                  className="w-full bg-slate-50 dark:bg-slate-900 border border-slate-200 dark:border-slate-700 text-xs text-slate-800 dark:text-white p-2.5 rounded-xl outline-none focus:border-[#117097] font-bold mt-1 resize-none placeholder-slate-400"
                  value={gastoForm.nota || ''} onChange={e => setGastoForm({...gastoForm, nota: e.target.value})}></textarea>
              </div>

              <div className="flex justify-end gap-3 pt-4 border-t border-slate-200 dark:border-slate-700">
                <button type="button" onClick={() => setModalAbierto(false)} className="px-5 py-2.5 bg-slate-200 dark:bg-slate-700 hover:bg-slate-300 dark:hover:bg-slate-600 text-slate-700 dark:text-slate-200 font-black text-xs uppercase tracking-wider rounded-xl transition-colors cursor-pointer">
                  Cancelar
                </button>
                <button type="submit" className="px-6 py-2.5 bg-[#117097] hover:bg-[#0a4c68] text-white font-black text-xs uppercase tracking-wider rounded-xl shadow-lg transition-colors cursor-pointer">
                  {gastoForm.id_editando ? "💾 Guardar Cambios" : "🚀 Registrar Egreso"}
                </button>
              </div>

            </form>

          </div>
        </div>
      )}

    </div>
  );
}