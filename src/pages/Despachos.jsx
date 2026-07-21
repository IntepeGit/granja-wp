import React, { useState, useEffect } from 'react';
import { jsPDF } from "jspdf";
import autoTable from "jspdf-autotable";
import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';

export default function Despachos({ 
  despachoForm, 
  setDespachoForm, 
  listaClientes, 
  listaInvernaderos, 
  actualizarFilaDespacho, 
  guardarDespachoCompleto, 
  datosDespachos,
  datosPagos,
  mostrarAlerta,
  guardarDespacho,
  eliminarDespacho, 
  prepararEdicion, 
  prepararEdicionDespacho,
  imprimirPDF, 
  cargarTodo,
  supabase,
  userRole 
}) {
  const [busqueda, setBusqueda] = useState('');
  const opcionesEscala = ["Kilo", "Bulto", "Caja", "Unidad", "Gramos", "Canastilla", "Amarres", "Libras"];

  const formatoPesos = (valor) => new Intl.NumberFormat('es-CO', { 
    style: 'currency', 
    currency: 'COP', 
    minimumFractionDigits: 0 
  }).format(valor || 0);

  // AUTOMATIZACIÓN DEL CONSECUTIVO DE REMISIÓN
  useEffect(() => {
    if (!despachoForm.id_editando && datosDespachos && datosDespachos.length > 0) {
      const numeros = datosDespachos.map(d => parseInt(d.numero_remision, 10)).filter(num => !isNaN(num));
      if (numeros.length > 0) {
        const maxRemision = Math.max(...numeros);
        const siguienteRemision = maxRemision + 1;
        if (String(despachoForm.numero_remision) !== String(siguienteRemision)) {
          setDespachoForm(prev => ({
            ...prev,
            numero_remision: String(siguienteRemision),
            errorDuplicado: false
          }));
        }
      }
    } else if (!despachoForm.id_editando && (!datosDespachos || datosDespachos.length === 0)) {
      if (!despachoForm.numero_remision) {
        setDespachoForm(prev => ({ ...prev, numero_remision: '1' }));
      }
    }
  }, [datosDespachos, despachoForm.id_editando]);

  // EXPORTACIÓN DE DESPACHOS A EXCEL EN 2 HOJAS
  const exportarDespachosAExcel = async () => {
    if (!datosDespachos || datosDespachos.length === 0) {
      if (mostrarAlerta) mostrarAlerta("No hay registros de despachos para exportar", "error");
      return;
    }

    try {
      const workbook = new ExcelJS.Workbook();

      // HOJA 1: RESUMEN GENERAL
      const wsResumen = workbook.addWorksheet('Resumen de Remisiones');
      wsResumen.columns = [
        { header: 'FECHA DESPACHO', key: 'fecha', width: 16 },
        { header: 'REMISIÓN N°', key: 'remision', width: 15 },
        { header: 'INVERNADERO ORIGEN', key: 'inv', width: 22 },
        { header: 'CLIENTE / DESTINO', key: 'cliente', width: 30 },
        { header: 'VALOR TOTAL CARGA', key: 'total', width: 22 }
      ];

      datosDespachos.forEach(d => {
        const numRem = d.numero_remision || 'S/N';
        const productosImpresos = new Set();
        let sumatoriaProductosReal = 0;

        (d.detalle_ventas || []).forEach(item => {
          const claveUnica = `${numRem}-${String(item.descripcion).trim().toUpperCase()}`;
          if (productosImpresos.has(claveUnica)) return;
          productosImpresos.add(claveUnica);

          const c = parseFloat(item.cantidad || 0);
          const p = parseFloat(item.precio_unitario || item.precio || 0);
          sumatoriaProductosReal += (c * p);
        });

        const valorFinalCarga = sumatoriaProductosReal > 0 ? sumatoriaProductosReal : parseFloat(d.total_venta || 0);

        wsResumen.addRow({
          fecha: d.fecha_venta ? String(d.fecha_venta).split('T')[0] : 'S/F',
          remision: d.numero_remision || 'S/N',
          inv: (d.invernaderos?.nombre || 'General').toUpperCase(),
          cliente: (d.clientes?.nombre_completo || 'Particular').toUpperCase(),
          total: valorFinalCarga
        });
      });

      const ultimaFilaResumen = wsResumen.rowCount;
      const filaTotalResumen = wsResumen.addRow({
        fecha: '', remision: '', inv: '',
        cliente: 'TOTAL GENERAL DESPACHOS:',
        total: { formula: `SUM(E2:E${ultimaFilaResumen})` }
      });

      wsResumen.getRow(1).height = 24;
      wsResumen.getRow(1).eachCell(c => {
        c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E293B' } };
        c.font = { name: 'Arial', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
        c.alignment = { vertical: 'middle', horizontal: 'center' };
      });

      wsResumen.eachRow((row, idx) => {
        if (idx === 1) return;
        row.height = 20;

        if (idx === wsResumen.rowCount) {
          row.eachCell(cell => {
            cell.border = { top: { style: 'thin' }, bottom: { style: 'double' } };
          });
          row.getCell('cliente').font = { name: 'Arial', size: 10, bold: true };
          row.getCell('total').font = { name: 'Arial', size: 11, bold: true, color: { argb: 'FF15803D' } };
          row.getCell('total').numFmt = '"$"#,##0';
          row.getCell('total').alignment = { horizontal: 'right', vertical: 'middle' };
          return;
        }

        const esCebra = idx % 2 === 0;
        row.eachCell((cell, colNum) => {
          cell.font = { name: 'Arial', size: 9 };
          if (esCebra) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF1F5F9' } };
          if (colNum === 5) {
            cell.numFmt = '"$"#,##0';
            cell.alignment = { horizontal: 'right', vertical: 'middle' };
          } else if (colNum === 1 || colNum === 2) {
            cell.alignment = { horizontal: 'center', vertical: 'middle' };
          } else {
            cell.alignment = { horizontal: 'left', vertical: 'middle' };
          }
        });
      });
      wsResumen.autoFilter = { from: { row: 1, column: 1 }, to: { row: ultimaFilaResumen, column: 5 } };

      // HOJA 2: DESGLOSE DE PRODUCTOS
      const wsDetalle = workbook.addWorksheet('Desglose de Productos');
      wsDetalle.columns = [
        { header: 'REMISIÓN N°', key: 'remision', width: 15 },
        { header: 'PRODUCTO / VARIEDAD', key: 'prod', width: 28 },
        { header: 'CANTIDAD', key: 'cant', width: 14 },
        { header: 'ESCALA', key: 'escala', width: 14 },
        { header: 'PRECIO UNITARIO', key: 'precio', width: 18 },
        { header: 'SUBTOTAL ITEM', key: 'subtotal', width: 20 }
      ];

      datosDespachos.forEach(d => {
        const numRem = d.numero_remision || 'S/N';
        const productosImpresos = new Set();

        (d.detalle_ventas || []).forEach(item => {
          const claveUnica = `${numRem}-${String(item.descripcion).trim().toUpperCase()}`;
          if (productosImpresos.has(claveUnica)) return;
          productosImpresos.add(claveUnica);

          const c = parseFloat(item.cantidad || 0);
          const p = parseFloat(item.precio_unitario || item.precio || 0);
          wsDetalle.addRow({
            remision: numRem,
            prod: String(item.descripcion || '').toUpperCase(),
            cant: c,
            escala: String(item.escala || 'Kilo').toUpperCase(),
            precio: p,
            subtotal: c * p
          });
        });
      });

      const ultimaFilaDetalle = wsDetalle.rowCount;
      const filaTotalDetalle = wsDetalle.addRow({
        remision: '', prod: '', cant: '', escala: '',
        precio: 'TOTAL CARGAS:',
        subtotal: { formula: `SUM(F2:F${ultimaFilaDetalle})` }
      });

      wsDetalle.getRow(1).height = 24;
      wsDetalle.getRow(1).eachCell(c => {
        c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF70AD47' } };
        c.font = { name: 'Arial', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
        c.alignment = { vertical: 'middle', horizontal: 'center' };
      });

      wsDetalle.eachRow((row, idx) => {
        if (idx === 1) return;
        row.height = 18;

        if (idx === wsDetalle.rowCount) {
          row.eachCell(cell => {
            cell.border = { top: { style: 'thin' }, bottom: { style: 'double' } };
          });
          row.getCell('precio').font = { name: 'Arial', size: 10, bold: true };
          row.getCell('subtotal').font = { name: 'Arial', size: 11, bold: true, color: { argb: 'FF15803D' } };
          row.getCell('subtotal').numFmt = '"$"#,##0';
          row.getCell('subtotal').alignment = { horizontal: 'right', vertical: 'middle' };
          return;
        }

        const esCebra = idx % 2 === 0;
        row.eachCell((cell, colNum) => {
          cell.font = { name: 'Arial', size: 9 };
          if (esCebra) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE2EFDA' } };
          
          if (colNum === 5 || colNum === 6) {
            cell.numFmt = '"$"#,##0';
            cell.alignment = { horizontal: 'right', vertical: 'middle' };
          } else if (colNum === 1 || colNum === 3 || colNum === 4) {
            cell.alignment = { horizontal: 'center', vertical: 'middle' };
          } else {
            cell.alignment = { horizontal: 'left', vertical: 'middle' };
          }
        });
      });
      wsDetalle.autoFilter = { from: { row: 1, column: 1 }, to: { row: ultimaFilaDetalle, column: 6 } };

      const buffer = await workbook.xlsx.writeBuffer();
      const fechaHoy = new Date().toISOString().split('T')[0];
      const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      saveAs(blob, `REPORTE_DESPACHOS_REMISIONES_${fechaHoy}.xlsx`);

      if (mostrarAlerta) mostrarAlerta("Despachos exportados a Excel con éxito", "exito");
    } catch (error) {
      console.error("Error al exportar Excel de Despachos:", error);
    }
  };

  const limpiarTextoParaPDF = (texto) => {
    if (!texto) return '';
    return String(texto)
      .replace(/[\uE000-\uF8FF]|\uD83C[\uDC00-\uDFFF]|\uD83D[\uDC00-\uDFFF]|[\u2011-\u26FF]|\uD83E[\uDD00-\uDFFF]/g, '')
      .replace(/\s+/g, ' ')
      .trim();
  };

  // IMPRESIÓN DE REMISIONES EN PDF
  const ejecutarImpresionDespachoLocal = (d) => {
    try {
      if (!d) return;
      const doc = new jsPDF({ orientation: 'portrait', unit: 'mm', format: [105, 148] });
      doc.setDrawColor(112, 173, 71); doc.setLineWidth(0.8); doc.rect(4, 4, 97, 140);
      try { doc.addImage('/Logopapel.png', 'PNG', 42.5, 6, 20, 20); } catch (e) {}

      const nRemision = d.numero_remision || 'S/N';
      doc.setFont("helvetica", "bold"); doc.setFontSize(8); doc.setTextColor(60, 60, 60);
      doc.text(`REMISIÓN N°: ${nRemision}`, 6, 11);

      doc.setFont("helvetica", "bold"); doc.setFontSize(11); doc.setTextColor(40, 80, 40);
      doc.text("REMISIÓN DE DESPACHO", 52.5, 29, { align: "center" });

      const yBase = 33;
      doc.setFillColor(242, 242, 242); doc.rect(6, yBase, 93, 5, 'F'); doc.rect(6, yBase + 10, 93, 5, 'F');
      doc.setDrawColor(210, 210, 210); doc.setLineWidth(0.2); doc.rect(6, yBase, 93, 15);
      doc.line(6, yBase + 5, 99, yBase + 5); doc.line(6, yBase + 10, 99, yBase + 10); doc.line(52, yBase, 52, yBase + 5);

      const clienteNom = d.clientes?.nombre_completo || 'PARTICULAR';
      const invernaderoNom = d.invernaderos?.nombre || 'GENERAL';

      doc.setFont("helvetica", "bold"); doc.setFontSize(7); doc.setTextColor(0);
      doc.text("Fecha Despacho:", 7, yBase + 3.5);
      doc.setFont("helvetica", "normal"); doc.text(String(d.fecha_venta || '').split('T')[0], 28, yBase + 3.5);
      doc.setFont("helvetica", "bold"); doc.text("Origen Bloque:", 53, yBase + 3.5);
      doc.setFont("helvetica", "normal"); doc.text(limpiarTextoParaPDF(invernaderoNom).toUpperCase(), 71, yBase + 3.5);
      doc.setFont("helvetica", "bold"); doc.text("Cliente / Destino:", 7, yBase + 8.5);
      doc.setFont("helvetica", "normal"); doc.text(limpiarTextoParaPDF(clienteNom).toUpperCase(), 28, yBase + 8.5);
      doc.setFont("helvetica", "bold"); doc.text("NIT / CC:", 7, yBase + 13.5);
      doc.setFont("helvetica", "normal"); doc.text(d.clientes?.nit_cc || 'N/A', 19, yBase + 13.5);

      const filasTabla = (d.detalle_ventas || []).map(item => [
        limpiarTextoParaPDF(item.descripcion).toUpperCase(),
        `${item.cantidad} ${item.escala || 'Kilo'}`,
        formatoPesos(item.precio_unitario || item.precio),
        formatoPesos(item.subtotal || (item.cantidad * (item.precio_unitario || item.precio)))
      ]);

      autoTable(doc, {
        startY: yBase + 18, margin: { left: 6, right: 6 },
        head: [["Descripción Producto", "Cantidad", "Precio Unit.", "Subtotal"]], body: filasTabla, theme: 'grid',
        styles: { font: 'helvetica', fontSize: 6.5, cellPadding: 1.5, lineWidth: 0.1, lineColor: [210, 210, 210] },
        headStyles: { fillColor: [112, 173, 71], textColor: [255, 255, 255], halign: 'center', fontStyle: 'bold' },
        columnStyles: { 0: { cellWidth: 40, halign: 'left' }, 1: { cellWidth: 18, halign: 'center' }, 2: { cellWidth: 17, halign: 'right' }, 3: { cellWidth: 18, halign: 'right' } }
      });

      const yTotal = doc.lastAutoTable.finalY + 6;
      doc.setFont("helvetica", "bold"); doc.setFontSize(8.5); doc.setTextColor(0);
      doc.text("TOTAL CARGA DESPACHADA:", 35, yTotal);
      doc.text(formatoPesos(d.total_venta), 99, yTotal, { align: 'right' });

      const yFirmas = 134;
      doc.setDrawColor(150); doc.setLineWidth(0.2);
      doc.line(8, yFirmas, 48, yFirmas); doc.line(56, yFirmas, 96, yFirmas);
      doc.setFont("helvetica", "bold"); doc.setFontSize(6.5); doc.setTextColor(50);
      doc.text("DESPACHADO POR (GRANJA)", 28, yFirmas + 3, { align: "center" });
      doc.text("RECIBIDO CONFORME (CLIENTE)", 76, yFirmas + 3, { align: "center" });

      doc.save(`REM_DESPACHO_N_${nRemision}_${clienteNom.replace(/ /g, "_")}.pdf`);
    } catch (error) {
      console.error("Error crítico de rendering jsPDF:", error);
    }
  };

  const totalFormulario = (despachoForm?.filas || []).reduce((acc, fila) => {
    return acc + (parseFloat(fila.cantidad || 0) * parseFloat(fila.precio || 0));
  }, 0);

  const agregarFila = () => {
    setDespachoForm({
      ...despachoForm,
      filas: [...despachoForm.filas, { producto: '', escala: '', cantidad: '', precio: '' }]
    });
  };

  const eliminarFilaFormulario = (index) => {
    if (despachoForm.filas.length === 1) return;
    const nuevasFilas = despachoForm.filas.filter((_, i) => i !== index);
    setDespachoForm({ ...despachoForm, filas: nuevasFilas });
  };

  // FILTRADO CON BARRA DE BÚSQUEDA INTEGRA EN PANTALLA
  const despachosFiltrados = (datosDespachos || [])
    .filter(d => {
      const q = busqueda.toLowerCase();
      const numRem = (d.numero_remision || '').toString().toLowerCase();
      const cliente = (d.clientes?.nombre_completo || '').toLowerCase();
      const inv = (d.invernaderos?.nombre || '').toLowerCase();
      return numRem.includes(q) || cliente.includes(q) || inv.includes(q);
    })
    .sort((a, b) => b.id - a.id);

  return (
    <div className="space-y-6 pb-20 text-slate-800">
      <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
        
        {/* COLUMNA IZQUIERDA: FORMULARIO DE DESPACHO */}
        <div className="bg-white p-6 rounded-3xl shadow-xl border-t-8 border-green-700 h-fit">
          <div className="flex justify-between items-center mb-5 border-b pb-3">
            <h3 className="font-black text-slate-800 uppercase text-xs italic">
              {despachoForm.id_editando ? '📝 Editar Remisión' : '🚚 Nueva Remisión'}
            </h3>
            <div className="text-right bg-green-50 px-3 py-1.5 rounded-xl border border-green-100">
              <p className="text-[9px] font-black text-green-600 uppercase italic leading-none">Total Carga</p>
              <p className="text-sm font-black text-green-700 mt-0.5">{formatoPesos(totalFormulario)}</p>
            </div>
          </div>
          
          <form onSubmit={guardarDespachoCompleto} className="space-y-4">
            <div>
              <label className="text-[10px] font-bold text-gray-400 uppercase px-1 italic">Número de Remisión</label>
              <input
                type="text"
                placeholder="N° Remisión"
                className={`w-full p-2.5 border-2 rounded-xl font-black text-sm outline-none transition-all ${
                  despachoForm.errorDuplicado 
                    ? 'border-red-400 bg-red-50 text-red-600 focus:border-red-500' 
                    : 'border-blue-100 bg-blue-50/50 text-blue-600 focus:border-blue-500'
                }`}
                value={despachoForm.numero_remision || ''}
                onChange={(e) => setDespachoForm({...despachoForm, numero_remision: e.target.value})}
              />
              {despachoForm.errorDuplicado && (
                <span className="text-[9px] font-black text-red-600 block mt-1 ml-1 animate-pulse">
                  ⚠️ ESTE NÚMERO YA EXISTE
                </span>
              )}
            </div>

            <div>
              <label className="text-[10px] font-bold text-gray-400 uppercase px-1 italic">Fecha Despacho</label>
              <input 
                type="date" 
                className="w-full border-2 p-2.5 rounded-xl font-bold text-sm bg-white outline-none focus:border-green-700" 
                value={despachoForm.fecha_venta || ''} 
                onChange={e => setDespachoForm({...despachoForm, fecha_venta: e.target.value})} 
                required 
              />
            </div>

            <div>
              <label className="text-[10px] font-bold text-gray-400 uppercase px-1 italic">Invernadero (Origen)</label>
              <select 
                className="w-full border-2 p-2.5 rounded-xl font-bold text-xs bg-white outline-none focus:border-green-700"
                value={despachoForm.invernadero_id || ''}
                onChange={e => setDespachoForm({...despachoForm, invernadero_id: e.target.value})}
                required
              >
                <option value="">Seleccione bloque...</option>
                {listaInvernaderos?.filter(i => i.activo !== false).map(inv => <option key={inv.id} value={inv.id}>{inv.nombre?.toUpperCase()}</option>)}
              </select>
            </div>

            <div>
              <label className="text-[10px] font-bold text-gray-400 uppercase px-1 italic">Cliente (Destino)</label>
              <select 
                className="w-full border-2 p-2.5 rounded-xl font-bold text-xs bg-white outline-none focus:border-green-700"
                value={despachoForm.cliente_id || ''}
                onChange={e => setDespachoForm({...despachoForm, cliente_id: e.target.value})}
                required
              >
                <option value="">Seleccione cliente...</option>
                {listaClientes?.filter(cli => cli.activo !== false).map(cli => <option key={cli.id} value={cli.id}>{cli.nombre_completo?.toUpperCase()}</option>)}
              </select>
            </div>

            <div className="pt-2 space-y-3 border-t">
              <p className="text-[10px] font-black text-slate-700 uppercase tracking-wider italic">📦 Desglose de Productos</p>
              
              {despachoForm.filas?.map((fila, index) => (
                <div key={index} className="bg-slate-50 p-3 rounded-2xl border border-slate-200 space-y-2 relative">
                  {despachoForm.filas.length > 1 && (
                    <button 
                      type="button" 
                      onClick={() => eliminarFilaFormulario(index)}
                      className="absolute top-2 right-3 text-red-400 hover:text-red-600 font-black text-xs"
                      title="Eliminar fila"
                    >
                      ✕
                    </button>
                  )}
                  
                  <div>
                    <label className="text-[8px] font-black text-slate-400 uppercase">Producto / Variedad</label>
                    <input 
                      placeholder="Ej: Tomate Larga Vida" 
                      className="w-full p-2 border bg-white rounded-xl font-bold text-xs uppercase outline-none focus:border-green-500" 
                      value={fila.producto || ''} 
                      onChange={e => actualizarFilaDespacho(index, 'producto', e.target.value)} 
                      required 
                    />
                  </div>

                  <div className="grid grid-cols-2 gap-2">
                    <div>
                      <label className="text-[8px] font-black text-slate-400 uppercase">Cantidad</label>
                      <input 
                        type="number" 
                        step="any"
                        placeholder="0" 
                        className="w-full p-2 border bg-white rounded-xl font-black text-xs text-center outline-none focus:border-green-500" 
                        value={fila.cantidad || ''} 
                        onChange={e => actualizarFilaDespacho(index, 'cantidad', e.target.value)} 
                        required 
                      />
                    </div>
                    <div>
                      <label className="text-[8px] font-black text-slate-400 uppercase">Escala</label>
                      <select 
                        className="w-full p-2 border bg-white rounded-xl font-bold text-xs outline-none focus:border-green-500" 
                        value={fila.escala || ''} 
                        onChange={e => actualizarFilaDespacho(index, 'escala', e.target.value)} 
                        required
                      >
                        <option value="">...</option>
                        {opcionesEscala.map(o => <option key={o} value={o}>{o}</option>)}
                      </select>
                    </div>
                  </div>

                  <div className="grid grid-cols-2 gap-2 items-center pt-1">
                    <div>
                      <label className="text-[8px] font-black text-slate-400 uppercase">Precio Unit.</label>
                      <input 
                        type="text" 
                        placeholder="$ 0" 
                        className="w-full p-1.5 border bg-white rounded-xl font-black text-green-700 text-xs text-center outline-none focus:border-green-500" 
                        value={formatoPesos(fila.precio)} 
                        onChange={e => actualizarFilaDespacho(index, 'precio', e.target.value.replace(/\D/g, ""))} 
                        required 
                      />
                    </div>
                    <div className="text-right pr-1">
                      <p className="text-[8px] font-black text-slate-400 uppercase">Subtotal</p>
                      <p className="font-black text-slate-800 text-xs mt-1">
                        {formatoPesos(parseFloat(fila.cantidad || 0) * parseFloat(fila.precio || 0))}
                      </p>
                    </div>
                  </div>
                </div>
              ))}
              
              <button 
                type="button" 
                onClick={agregarFila}
                className="w-full bg-blue-50 text-blue-600 text-[10px] font-black py-2 rounded-xl border border-dashed border-blue-300 hover:bg-blue-100 transition-colors uppercase tracking-widest"
              >
                + Añadir Producto
              </button>
            </div>

            <button 
              type="submit" 
              className="w-full bg-green-700 text-white font-black py-3.5 rounded-xl shadow-md uppercase tracking-wider text-xs hover:bg-green-800 transition-all cursor-pointer"
            >
              {despachoForm.id_editando ? '💾 Actualizar Remisión' : '🚀 Guardar Remisión'}
            </button>
          </form>
        </div>

        {/* COLUMNA DERECHA: TABLA HISTÓRICA CON BÚSQUEDA E IMPRESIÓN */}
        <div className="lg:col-span-2 bg-white rounded-3xl shadow-xl overflow-hidden border border-gray-200">
          
          <div className="p-4 bg-slate-800 text-white font-black text-xs uppercase tracking-widest italic flex flex-col sm:flex-row justify-between items-center gap-2">
            <span>Historial Reciente de Cargas ({despachosFiltrados.length})</span>
            
            <div className="flex items-center gap-2 flex-wrap w-full sm:w-auto justify-end">
              <button
                onClick={exportarDespachosAExcel}
                className="px-3 py-1.5 bg-emerald-600 hover:bg-emerald-700 text-white font-black rounded-xl shadow-md transition-all flex items-center gap-1 text-[10px] uppercase tracking-wider cursor-pointer"
              >
                📊 EXPORTAR A EXCEL
              </button>

              <input 
                type="text" 
                placeholder="🔍 Buscar remisión, cliente..." 
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
                  <th className="p-3.5 border-b border-gray-300">Fecha</th>
                  <th className="p-3.5 border-b border-gray-300 text-center">N° Remisión</th>
                  <th className="p-3.5 border-b border-gray-300">Cliente / Destino</th>
                  <th className="p-3.5 border-b border-gray-300">Productos (Cant + Escala)</th>
                  <th className="p-3.5 border-b border-gray-300 text-right">Total Carga</th>
                  <th className="p-3.5 border-b border-gray-300 text-center">Acciones</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-gray-200 font-bold text-slate-700">
                {despachosFiltrados.length === 0 ? (
                  <tr>
                    <td colSpan="6" className="p-8 text-center text-gray-400 italic font-bold">No hay registros de despachos coincidentes.</td>
                  </tr>
                ) : (
                  despachosFiltrados.map((d, index) => (
                    <tr 
                      key={d.id} 
                      className={`${index % 2 === 0 ? 'bg-white' : 'bg-slate-50'} hover:bg-sky-50 transition-colors border-l-4 border-green-700`}
                    >
                      <td className="p-3.5 font-bold text-slate-600 italic whitespace-nowrap">{d.fecha_venta}</td>
                      <td className="p-3.5 text-center font-black text-slate-900 text-sm">{d.numero_remision}</td>
                      <td className="p-3.5 uppercase font-bold text-slate-800">
                        <p className="font-black text-slate-900">{d.clientes?.nombre_completo || 'N/A'}</p>
                        <p className="text-[9px] text-[#117097] lowercase italic font-bold mt-0.5">🌿 Bloque: {d.invernaderos?.nombre || 'General'}</p>
                      </td>
                      <td className="p-3.5">
                        <div className="space-y-1">
                          {d.detalle_ventas?.map((item, i) => (
                            <div key={i} className="flex gap-1.5 items-center">
                              <p className="font-black text-slate-800 uppercase text-[10px] leading-none">{item.descripcion}</p>
                              <span className="bg-green-100 text-green-800 px-1.5 py-0.5 rounded-md text-[8px] font-black italic">
                                {item.cantidad} {item.escala}
                              </span>
                            </div>
                          ))}
                        </div>
                      </td>
                      <td className="p-3.5 text-right font-black text-green-800 text-xs whitespace-nowrap">
                        {formatoPesos(d.total_venta)}
                      </td>
                      <td className="p-3.5 text-center">
                        <div className="flex gap-1.5 justify-center">
                          <button 
                            onClick={() => ejecutarImpresionDespachoLocal(d)} 
                            className="px-2 py-1 bg-slate-800 text-white rounded-lg hover:bg-black transition-colors flex items-center gap-1 border border-slate-900 shadow"
                            title="Imprimir PDF"
                          >
                            <span className="text-[10px]">🖨️</span><span className="text-[9px] font-black tracking-wider">PDF</span>
                          </button>
                          
                          {userRole === 'admin' && (
                            <button 
                              onClick={() => eliminarDespacho(d.id)}
                              className="p-1 bg-red-100 text-red-700 rounded-lg hover:bg-red-600 hover:text-white transition-all border border-red-200 text-xs"
                              title="Eliminar Remisión"
                            >
                              🗑️
                            </button>
                          )}
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