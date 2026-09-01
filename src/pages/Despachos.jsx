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
  const [modalAbierto, setModalAbierto] = useState(false);
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

  // EXPORTACIÓN DE DESPACHOS A EXCEL (TOTALES CALCULADOS POR REMISIÓN Y FECHA)
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
        { header: 'REMISIÓN N°', key: 'remision', width: 15 },
        { header: 'FECHA DESPACHO', key: 'fecha', width: 16 },
        { header: 'INVERNADERO ORIGEN', key: 'inv', width: 22 },
        { header: 'CLIENTE / DESTINO', key: 'cliente', width: 30 },
        { header: 'VALOR TOTAL CARGA', key: 'total', width: 22 }
      ];

      datosDespachos.forEach(d => {
        const numRem = d.numero_remision || 'S/N';
        const fechaDespacho = d.fecha_venta ? String(d.fecha_venta).split('T')[0] : 'S/F';
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
          remision: numRem,
          fecha: fechaDespacho,
          inv: (d.invernaderos?.nombre || 'General').toUpperCase(),
          cliente: (d.clientes?.nombre_completo || 'Particular').toUpperCase(),
          total: valorFinalCarga
        });
      });

      const ultimaFilaResumen = wsResumen.rowCount;
      wsResumen.addRow({
        remision: '', fecha: '', inv: '',
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

      // HOJA 2: DESGLOSE DE PRODUCTOS (Agrupación precisa por Remisión + Fecha)
      const wsDetalle = workbook.addWorksheet('Desglose de Productos');
      wsDetalle.columns = [
        { header: 'REMISIÓN N°', key: 'remision', width: 15 },
        { header: 'FECHA DESPACHO', key: 'fecha', width: 15 },
        { header: 'INVERNADERO ORIGEN', key: 'inv', width: 20 },
        { header: 'CLIENTE / DESTINO', key: 'cliente', width: 28 },
        { header: 'PRODUCTO / VARIEDAD', key: 'prod', width: 25 },
        { header: 'CANTIDAD', key: 'cant', width: 12 },
        { header: 'ESCALA', key: 'escala', width: 12 },
        { header: 'PRECIO UNITARIO', key: 'precio', width: 18 },
        { header: 'SUBTOTAL ITEM', key: 'subtotal', width: 18 },
        { header: 'TOTAL x Remision', key: 'totalRemisionCol', width: 22 }
      ];

      const totalesPorRemisionFecha = {};
      datosDespachos.forEach(d => {
        const numRem = d.numero_remision || 'S/N';
        const fechaDespacho = d.fecha_venta ? String(d.fecha_venta).split('T')[0] : 'S/F';
        const llaveUnicaRemision = `${numRem}_${fechaDespacho}`;

        const productosImpresos = new Set();
        let sumatoria = 0;

        (d.detalle_ventas || []).forEach(item => {
          const claveUnica = `${numRem}-${String(item.descripcion).trim().toUpperCase()}`;
          if (productosImpresos.has(claveUnica)) return;
          productosImpresos.add(claveUnica);

          const c = parseFloat(item.cantidad || 0);
          const p = parseFloat(item.precio_unitario || item.precio || 0);
          sumatoria += (c * p);
        });

        totalesPorRemisionFecha[llaveUnicaRemision] = sumatoria > 0 ? sumatoria : parseFloat(d.total_venta || 0);
      });

      const filasFinalesRemision = [];
      let filaActualTracker = 2;

      datosDespachos.forEach(d => {
        const numRem = d.numero_remision || 'S/N';
        const fechaDespacho = d.fecha_venta ? String(d.fecha_venta).split('T')[0] : 'S/F';
        const llaveUnicaRemision = `${numRem}_${fechaDespacho}`;
        const invernaderoNom = (d.invernaderos?.nombre || 'General').toUpperCase();
        const clienteNom = (d.clientes?.nombre_completo || 'Particular').toUpperCase();
        const totalRemisionValor = totalesPorRemisionFecha[llaveUnicaRemision] || 0;

        const productosImpresos = new Set();
        let itemsDeEstaRemision = [];

        (d.detalle_ventas || []).forEach(item => {
          const claveUnica = `${numRem}-${String(item.descripcion).trim().toUpperCase()}`;
          if (productosImpresos.has(claveUnica)) return;
          productosImpresos.add(claveUnica);

          const c = parseFloat(item.cantidad || 0);
          const p = parseFloat(item.precio_unitario || item.precio || 0);

          itemsDeEstaRemision.push({
            c, p, desc: String(item.descripcion || '').toUpperCase(), escala: String(item.escala || 'Kilo').toUpperCase()
          });
        });

        itemsDeEstaRemision.forEach((it, idxItem) => {
          const esUltimoItem = idxItem === itemsDeEstaRemision.length - 1;

          wsDetalle.addRow({
            remision: numRem,
            fecha: fechaDespacho,
            inv: invernaderoNom,
            cliente: clienteNom,
            prod: it.desc,
            cant: it.c,
            escala: it.escala,
            precio: it.p,
            subtotal: it.c * it.p,
            totalRemisionCol: esUltimoItem ? totalRemisionValor : ''
          });
        });

        filaActualTracker += itemsDeEstaRemision.length;
        filasFinalesRemision.push(filaActualTracker - 1);
      });

      const ultimaFilaDetalle = wsDetalle.rowCount;
      wsDetalle.addRow({
        remision: '', fecha: '', inv: '', cliente: '', prod: '', cant: '', escala: '',
        precio: 'TOTAL CARGAS:',
        subtotal: { formula: `SUM(I2:I${ultimaFilaDetalle})` },
        totalRemisionCol: ''
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
          
          if (colNum === 8 || colNum === 9 || colNum === 10) {
            cell.numFmt = '"$"#,##0';
            cell.alignment = { horizontal: 'right', vertical: 'middle' };
          } else if (colNum === 1 || colNum === 2 || colNum === 6 || colNum === 7) {
            cell.alignment = { horizontal: 'center', vertical: 'middle' };
          } else {
            cell.alignment = { horizontal: 'left', vertical: 'middle' };
          }
        });

        if (filasFinalesRemision.includes(idx)) {
          for (let c = 1; c <= 10; c++) {
            const cellActual = row.getCell(c);
            cellActual.border = {
              ...(cellActual.border || {}),
              bottom: { style: 'double', color: { argb: 'FF000000' } }
            };
          }
          const cellTotalRemision = row.getCell(10);
          cellTotalRemision.font = { name: 'Arial', size: 9, bold: true, color: { argb: 'FF15803D' } };
        }
      });

      wsDetalle.autoFilter = { from: { row: 1, column: 1 }, to: { row: ultimaFilaDetalle, column: 10 } };

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

  const abrirModalParaNuevo = () => {
    setDespachoForm(prev => ({
      ...prev,
      id_editando: null,
      cliente_id: '',
      invernadero_id: '',
      filas: [{ producto: '', escala: '', cantidad: '', precio: '' }]
    }));
    setModalAbierto(true);
  };

  const handleGuardarWrapper = async (e) => {
    e.preventDefault();
    await guardarDespachoCompleto(e);
    setModalAbierto(false);
  };

  const despachosFiltrados = (datosDespachos || [])
    .filter(d => {
      const q = busqueda.toLowerCase();
      const numRem = (d.numero_remision || '').toString().toLowerCase();
      const cliente = (d.clientes?.nombre_completo || '').toLowerCase();
      const inv = (d.invernaderos?.nombre || '').toLowerCase();
      return numRem.includes(q) || cliente.includes(q) || inv.includes(q);
    })
    .sort((a, b) => b.id - b.id);

  return (
    <div className="space-y-6 pb-20 text-slate-800 dark:text-slate-200 font-sans transition-colors duration-300">
      
      {/* 🚀 CABECERA PRINCIPAL MODERNA */}
      <div className="bg-white dark:bg-slate-800 p-6 rounded-3xl shadow-xl border border-slate-200 dark:border-slate-700 flex flex-col md:flex-row justify-between items-center gap-4 transition-colors duration-300">
        <div className="space-y-1">
          <div className="flex items-center gap-3">
            <span className="text-2xl p-2 bg-green-700/10 dark:bg-emerald-500/20 rounded-xl text-green-700 dark:text-emerald-400">🚚</span>
            <h2 className="text-xl font-black tracking-tight text-slate-900 dark:text-white uppercase">Despachos y Remisiones</h2>
          </div>
          <p className="text-xs text-slate-500 dark:text-slate-400 font-medium">Registro de cargas, salidas a clientes y control de remisiones de venta.</p>
        </div>
        
        <div className="flex items-center gap-3 w-full md:w-auto">
          <button onClick={exportarDespachosAExcel} className="flex-1 md:flex-none px-4 py-2.5 bg-emerald-600 hover:bg-emerald-700 text-white font-black text-xs uppercase tracking-wider rounded-xl shadow-md transition-all flex items-center justify-center gap-2 cursor-pointer">
            <span>📊</span> Exportar Excel
          </button>
          <button onClick={abrirModalParaNuevo} className="flex-1 md:flex-none px-5 py-2.5 bg-green-700 hover:bg-green-800 text-white font-black text-xs uppercase tracking-wider rounded-xl shadow-md transition-all flex items-center justify-center gap-2 cursor-pointer">
            <span>➕</span> Nueva Remisión
          </button>
        </div>
      </div>

      {/* 🔍 BARRA DE BÚSQUEDA Y TOTALES */}
      <div className="bg-white dark:bg-slate-800 p-4 rounded-2xl shadow-md border border-slate-200 dark:border-slate-700 flex flex-col md:flex-row justify-between items-center gap-4 transition-colors duration-300">
        <div className="flex items-center gap-3">
          <span className="px-4 py-2 bg-slate-100 dark:bg-slate-700 text-slate-700 dark:text-slate-300 font-black text-[10px] uppercase rounded-xl border border-slate-200 dark:border-slate-600">
            Total Cargas: {despachosFiltrados.length}
          </span>
          <span className="px-3 py-1.5 bg-green-100 dark:bg-emerald-950/60 text-green-800 dark:text-emerald-400 font-black text-[10px] uppercase rounded-lg border border-green-200 dark:border-emerald-800">
            Logística
          </span>
        </div>
        
        <div className="relative w-full md:w-96">
          <span className="absolute inset-y-0 left-0 flex items-center pl-3 text-slate-400">🔍</span>
          <input 
            type="text" 
            placeholder="Buscar remisión, cliente, bloque..." 
            className="w-full bg-slate-50 dark:bg-slate-700/50 border border-slate-200 dark:border-slate-600 text-xs text-slate-800 dark:text-white rounded-xl pl-10 pr-4 py-2.5 outline-none focus:border-green-700 dark:focus:border-emerald-500 font-bold placeholder-slate-400"
            value={busqueda}
            onChange={e => setBusqueda(e.target.value)}
          />
        </div>
      </div>

      {/* 📊 TABLA DE DESPACHOS COMPRIMIDA EN DOBLE COLUMNA */}
      <div className="bg-white dark:bg-slate-800 rounded-3xl shadow-xl overflow-hidden border border-gray-200 dark:border-slate-700 transition-colors duration-300">
        <div className="overflow-x-auto max-h-[600px] overflow-y-auto">
          <table className="w-full text-left text-xs border-collapse">
           <thead>
              <tr className="bg-slate-100 dark:bg-slate-700/80 text-slate-700 dark:text-slate-300 uppercase font-black text-[10px] tracking-wider border-b border-slate-200 dark:border-slate-700 sticky top-0 z-10">
                <th className="py-2.5 px-3">Fecha</th>
                <th className="py-2.5 px-3 text-center">N° Remisión</th>
                <th className="py-2.5 px-3">Cliente / Destino</th>
                <th className="py-2.5 px-3 text-center">
                  PRODUCTOS (CANT + ESCALA)
                </th>
                <th className="py-2.5 px-3 text-right">Total Carga</th>
                <th className="py-2.5 px-3 text-center">Acciones</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100 dark:divide-slate-700/60 font-bold text-slate-700 dark:text-slate-300">
              {despachosFiltrados.length === 0 ? (
                <tr>
                  <td colSpan="6" className="p-6 text-center text-slate-400 dark:text-slate-500 italic font-bold">No hay registros de despachos coincidentes.</td>
                </tr>
              ) : (
                despachosFiltrados.map((d, index) => (
                  <tr 
                    key={d.id} 
                    className={`${index % 2 === 0 ? 'bg-white dark:bg-slate-800' : 'bg-slate-50/50 dark:bg-slate-800/60'} hover:bg-sky-50/50 dark:hover:bg-slate-700/50 transition-colors border-l-8 border-green-700 dark:border-emerald-600`}
                  >
                    <td className="py-2 px-3 font-black text-slate-900 dark:text-white whitespace-nowrap">{d.fecha_venta}</td>
                    
                    <td className="py-2 px-3 text-center whitespace-nowrap">
                      <span className="bg-slate-800 dark:bg-slate-900 text-white px-2 py-0.5 rounded-md text-[10px] font-black shadow-sm">
                        {d.numero_remision}
                      </span>
                    </td>
                    
                    <td className="py-2 px-3 uppercase font-bold text-slate-800 dark:text-slate-200">
                      <p className="font-black text-slate-900 dark:text-white text-xs">{d.clientes?.nombre_completo || 'N/A'}</p>
                      <p className="text-[9px] text-[#117097] dark:text-sky-400 lowercase italic font-bold mt-0.5">🌿 Bloque: {d.invernaderos?.nombre || 'General'}</p>
                    </td>
                    
                    {/* ⚡ PRODUCTOS EN DOBLE COLUMNA CENTRADOS */}
                    <td className="py-2 px-3">
                      <div className="grid grid-cols-2 gap-x-4 gap-y-1 max-w-md mx-auto">
                        {d.detalle_ventas?.map((item, i) => (
                          <div key={i} className="flex gap-1.5 items-center justify-between bg-slate-100 dark:bg-slate-900/50 px-2 py-1 rounded border border-slate-200 dark:border-slate-700/50">
                            <span className="font-black text-slate-700 dark:text-slate-300 uppercase text-[9px] truncate" title={item.descripcion}>{item.descripcion}</span>
                            <span className="bg-green-100 dark:bg-emerald-950/80 text-green-800 dark:text-emerald-400 border border-green-200 dark:border-emerald-800 px-1 py-0.5 rounded text-[8px] font-black italic whitespace-nowrap">
                              {item.cantidad} {item.escala}
                            </span>
                          </div>
                        ))}
                      </div>
                    </td>
                    
                    <td className="py-2 px-3 text-right font-black text-green-800 dark:text-emerald-400 text-xs whitespace-nowrap">
                      {formatoPesos(d.total_venta)}
                    </td>
                    
                    <td className="py-2 px-3 text-center whitespace-nowrap">
                      <div className="flex gap-1 justify-center">
                        <button 
                          onClick={() => ejecutarImpresionDespachoLocal(d)} 
                          className="px-2 py-1 bg-slate-800 dark:bg-slate-700 text-white rounded-lg hover:bg-black dark:hover:bg-slate-600 transition-colors flex items-center gap-1 border border-slate-900 dark:border-slate-600 shadow-sm cursor-pointer"
                          title="Imprimir PDF"
                        >
                          <span className="text-[10px]">🖨️</span><span className="text-[8px] font-black tracking-wider">PDF</span>
                        </button>
                        
                        {userRole === 'admin' && (
                          <button 
                            onClick={() => eliminarDespacho(d.id)}
                            className="p-1 bg-red-100 dark:bg-red-950/60 text-red-700 dark:text-red-400 rounded-lg hover:bg-red-700 hover:text-white transition-all border border-red-200 dark:border-red-900 text-[10px] cursor-pointer"
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

      {/* 🧊 MODAL FLOTANTE PARA NUEVO / EDITAR DESPACHO */}
      {modalAbierto && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 backdrop-blur-sm p-4 overflow-y-auto">
          <div className="bg-white dark:bg-slate-800 border border-slate-200 dark:border-slate-700 w-full max-w-2xl rounded-3xl shadow-2xl overflow-hidden animate-in fade-in zoom-in-95 duration-200 my-8 text-slate-800 dark:text-slate-200">
            
            <div className="p-5 bg-slate-900 text-white border-b border-slate-800 flex justify-between items-center">
              <div className="flex items-center gap-3">
                <span className="text-xl">🚚</span>
                <div>
                  <h3 className="text-base font-black uppercase tracking-tight">
                    {despachoForm.id_editando ? 'Editar Remisión' : 'Nueva Remisión'}
                  </h3>
                  <p className="text-[10px] text-slate-400 font-medium">Registro de despacho y control de carga.</p>
                </div>
              </div>
              <button onClick={() => setModalAbierto(false)} className="text-slate-400 hover:text-white text-lg font-black px-3 py-1 rounded-xl bg-slate-800 cursor-pointer">✕</button>
            </div>

            <form onSubmit={handleGuardarWrapper} className="p-6 space-y-4 max-h-[80vh] overflow-y-auto">
              
              <div className="grid grid-cols-2 gap-4">
                <div>
                  <label className="text-[10px] font-black text-slate-500 dark:text-slate-400 uppercase tracking-wider">N° Remisión</label>
                  <input
                    type="text"
                    placeholder="N° Remisión"
                    className={`w-full p-2.5 border-2 rounded-xl font-black text-sm outline-none transition-all mt-1 ${
                      despachoForm.errorDuplicado 
                        ? 'border-red-400 bg-red-50 dark:bg-red-950/50 text-red-600 dark:text-red-400 focus:border-red-500' 
                        : 'border-slate-200 dark:border-slate-700 bg-slate-50 dark:bg-slate-900 text-slate-800 dark:text-white focus:border-green-700'
                    }`}
                    value={despachoForm.numero_remision || ''}
                    onChange={(e) => setDespachoForm({...despachoForm, numero_remision: e.target.value})}
                  />
                  {despachoForm.errorDuplicado && (
                    <span className="text-[9px] font-black text-red-600 dark:text-red-400 block mt-1 animate-pulse">
                      ⚠️ ESTE NÚMERO YA EXISTE
                    </span>
                  )}
                </div>
                <div>
                  <label className="text-[10px] font-black text-slate-500 dark:text-slate-400 uppercase tracking-wider">Fecha Despacho *</label>
                  <input 
                    type="date" 
                    className="w-full bg-slate-50 dark:bg-slate-900 border border-slate-200 dark:border-slate-700 text-slate-800 dark:text-white p-2.5 rounded-xl font-bold text-sm outline-none focus:border-green-700 mt-1" 
                    value={despachoForm.fecha_venta || ''} 
                    onChange={e => setDespachoForm({...despachoForm, fecha_venta: e.target.value})} 
                    required 
                  />
                </div>
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                <div>
                  <label className="text-[10px] font-black text-slate-500 dark:text-slate-400 uppercase tracking-wider">Invernadero (Origen) *</label>
                  <select 
                    className="w-full bg-slate-50 dark:bg-slate-900 border border-slate-200 dark:border-slate-700 text-slate-800 dark:text-white p-2.5 rounded-xl font-bold text-xs uppercase outline-none focus:border-green-700 mt-1"
                    value={despachoForm.invernadero_id || ''}
                    onChange={e => setDespachoForm({...despachoForm, invernadero_id: e.target.value})}
                    required
                  >
                    <option value="">Seleccione bloque...</option>
                    {listaInvernaderos?.filter(i => i.activo !== false).map(inv => <option key={inv.id} value={inv.id}>{inv.nombre?.toUpperCase()}</option>)}
                  </select>
                </div>
                <div>
                  <label className="text-[10px] font-black text-slate-500 dark:text-slate-400 uppercase tracking-wider">Cliente (Destino) *</label>
                  <select 
                    className="w-full bg-slate-50 dark:bg-slate-900 border border-slate-200 dark:border-slate-700 text-slate-800 dark:text-white p-2.5 rounded-xl font-bold text-xs uppercase outline-none focus:border-green-700 mt-1"
                    value={despachoForm.cliente_id || ''}
                    onChange={e => setDespachoForm({...despachoForm, cliente_id: e.target.value})}
                    required
                  >
                    <option value="">Seleccione cliente...</option>
                    {listaClientes?.filter(cli => cli.activo !== false).map(cli => <option key={cli.id} value={cli.id}>{cli.nombre_completo?.toUpperCase()}</option>)}
                  </select>
                </div>
              </div>

              <div className="pt-2 space-y-3 border-t border-slate-200 dark:border-slate-700">
                <div className="flex justify-between items-center">
                  <p className="text-[10px] font-black text-slate-500 dark:text-slate-400 uppercase tracking-wider">📦 Desglose de Productos</p>
                  <div className="text-right bg-green-50 dark:bg-emerald-950/60 px-3 py-1.5 rounded-xl border border-green-200 dark:border-emerald-800 shadow-inner">
                    <p className="text-[8px] font-black text-green-700 dark:text-emerald-400 uppercase tracking-wider leading-none">Total Carga</p>
                    <p className="text-sm font-black text-green-800 dark:text-emerald-300 mt-0.5">{formatoPesos(totalFormulario)}</p>
                  </div>
                </div>
                
                {despachoForm.filas?.map((fila, index) => (
                  <div key={index} className="bg-slate-50 dark:bg-slate-900/60 p-4 rounded-2xl border border-slate-200 dark:border-slate-700 space-y-3 relative shadow-sm">
                    {despachoForm.filas.length > 1 && (
                      <button 
                        type="button" 
                        onClick={() => eliminarFilaFormulario(index)}
                        className="absolute top-2 right-2 text-red-400 hover:text-red-600 bg-red-50 dark:bg-red-950/50 hover:bg-red-100 rounded-lg p-1.5 transition-colors cursor-pointer"
                        title="Eliminar fila"
                      >
                        ✕
                      </button>
                    )}
                    
                    <div>
                      <label className="text-[9px] font-black text-slate-500 dark:text-slate-400 uppercase tracking-wider">Producto / Variedad *</label>
                      <input 
                        placeholder="Ej: Tomate Larga Vida" 
                        className="w-full p-2.5 border border-slate-200 dark:border-slate-700 bg-white dark:bg-slate-800 text-slate-800 dark:text-white rounded-xl font-bold text-xs uppercase outline-none focus:border-green-500 mt-1 placeholder-slate-400" 
                        value={fila.producto || ''} 
                        onChange={e => actualizarFilaDespacho(index, 'producto', e.target.value)} 
                        required 
                      />
                    </div>

                    <div className="grid grid-cols-2 gap-4">
                      <div>
                        <label className="text-[9px] font-black text-slate-500 dark:text-slate-400 uppercase tracking-wider">Cantidad *</label>
                        <input 
                          type="number" 
                          step="any"
                          placeholder="0" 
                          className="w-full p-2.5 border border-slate-200 dark:border-slate-700 bg-white dark:bg-slate-800 text-slate-800 dark:text-white rounded-xl font-black text-sm text-center outline-none focus:border-green-500 mt-1 placeholder-slate-400" 
                          value={fila.cantidad || ''} 
                          onChange={e => actualizarFilaDespacho(index, 'cantidad', e.target.value)} 
                          required 
                        />
                      </div>
                      <div>
                        <label className="text-[9px] font-black text-slate-500 dark:text-slate-400 uppercase tracking-wider">Escala *</label>
                        <select 
                          className="w-full p-2.5 border border-slate-200 dark:border-slate-700 bg-white dark:bg-slate-800 text-slate-800 dark:text-white rounded-xl font-bold text-xs uppercase outline-none focus:border-green-500 mt-1" 
                          value={fila.escala || ''} 
                          onChange={e => actualizarFilaDespacho(index, 'escala', e.target.value)} 
                          required
                        >
                          <option value="">...</option>
                          {opcionesEscala.map(o => <option key={o} value={o}>{o}</option>)}
                        </select>
                      </div>
                    </div>

                    <div className="grid grid-cols-2 gap-4 items-center pt-2 border-t border-slate-100 dark:border-slate-800">
                      <div>
                        <label className="text-[9px] font-black text-slate-500 dark:text-slate-400 uppercase tracking-wider">Precio Unitario ($) *</label>
                        <input 
                          type="text" 
                          placeholder="$ 0" 
                          className="w-full p-2 border-2 border-green-200 dark:border-emerald-800 bg-white dark:bg-slate-800 rounded-xl font-black text-green-800 dark:text-emerald-400 text-sm text-center outline-none focus:border-green-600 mt-1" 
                          value={formatoPesos(fila.precio)} 
                          onChange={e => actualizarFilaDespacho(index, 'precio', e.target.value.replace(/\D/g, ""))} 
                          required 
                        />
                      </div>
                      <div className="text-right">
                        <p className="text-[9px] font-black text-slate-500 dark:text-slate-400 uppercase tracking-wider">Subtotal Item</p>
                        <p className="font-black text-slate-800 dark:text-white text-base mt-1">
                          {formatoPesos(parseFloat(fila.cantidad || 0) * parseFloat(fila.precio || 0))}
                        </p>
                      </div>
                    </div>
                  </div>
                ))}
                
                <button 
                  type="button" 
                  onClick={agregarFila}
                  className="w-full bg-blue-50 dark:bg-sky-950/40 text-blue-700 dark:text-sky-400 text-[10px] font-black py-2.5 rounded-xl border border-dashed border-blue-300 dark:border-sky-800 hover:bg-blue-100 dark:hover:bg-sky-900/40 transition-colors uppercase tracking-widest cursor-pointer"
                >
                  + Añadir Otro Producto
                </button>
              </div>

              <div className="flex justify-end gap-3 pt-4 border-t border-slate-200 dark:border-slate-700">
                <button type="button" onClick={() => setModalAbierto(false)} className="px-5 py-2.5 bg-slate-200 dark:bg-slate-700 hover:bg-slate-300 dark:hover:bg-slate-600 text-slate-700 dark:text-slate-200 font-black text-xs uppercase tracking-wider rounded-xl transition-colors cursor-pointer">
                  Cancelar
                </button>
                <button type="submit" className="px-6 py-2.5 bg-green-700 hover:bg-green-800 text-white font-black text-xs uppercase tracking-wider rounded-xl shadow-lg transition-colors cursor-pointer">
                  {despachoForm.id_editando ? '💾 Actualizar Remisión' : '🚀 Guardar Remisión'}
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

    </div>
  );
}