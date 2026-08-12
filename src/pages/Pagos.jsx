import React, { useState, useEffect } from 'react';
import { supabase } from '../lib/supabase';
import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';
import { jsPDF } from "jspdf";
import autoTable from "jspdf-autotable";

export default function Pagos({ 
  pagoForm, setPagoForm, listaClientes, datosDespachos, 
  datosPagos, mostrarAlerta, cargarTodo, 
  guardarPago, prepararEdicionPago, eliminarPago 
}) {

  // Estados locales para el modal y filtros
  const [modalAbierto, setModalAbierto] = useState(false);
  const [filtroClienteId, setFiltroClienteId] = useState('');
  const [filtroDespachoId, setFiltroDespachoId] = useState('');

  // Estados locales para el medio de pago
  const [modoMedio, setModoMedio] = useState('Efectivo');
  const [bancoPersonalizado, setBancoPersonalizado] = useState('');

  const listaMediosPredeterminados = [
    'Efectivo',
    'Bre-B',
    'Transferencia Bancolombia',
    'Nequi / Daviplata',
    'Consignación Bancaria',
    'Cheque',
    'OTRO_MANUAL'
  ];

  // Sincronizar el estado local cuando se edita un abono o cambia el formulario
  useEffect(() => {
    if (pagoForm.medio_pago) {
      if (listaMediosPredeterminados.includes(pagoForm.medio_pago)) {
        setModoMedio(pagoForm.medio_pago);
        setBancoPersonalizado('');
      } else {
        setModoMedio('OTRO_MANUAL');
        setBancoPersonalizado(pagoForm.medio_pago);
      }
    } else {
      setModoMedio('Efectivo');
    }
  }, [pagoForm.id_editando, pagoForm.despacho_id]);

  // Si se selecciona un cliente/remisión en el filtro, actualizamos el formulario base
  useEffect(() => {
    if (filtroClienteId && filtroDespachoId && !pagoForm.id_editando) {
      setPagoForm(prev => ({
        ...prev,
        cliente_id: filtroClienteId,
        despacho_id: filtroDespachoId
      }));
    }
  }, [filtroClienteId, filtroDespachoId]);

  const formatoPesos = (valor) => new Intl.NumberFormat('es-CO', { 
    style: 'currency', 
    currency: 'COP', 
    minimumFractionDigits: 0 
  }).format(valor || 0);

  const formatearMascaraMoneda = (valorRaw) => {
    const numeroLimpio = String(valorRaw).replace(/\D/g, "");
    if (!numeroLimpio) return "";
    return new Intl.NumberFormat('es-CO', {
      style: 'currency',
      currency: 'COP',
      minimumFractionDigits: 0
    }).format(numeroLimpio);
  };

  const alCambiarMedioSelect = (val) => {
    setModoMedio(val);
    if (val === 'OTRO_MANUAL') {
      const valorFinal = bancoPersonalizado.toUpperCase().trim() || 'OTRO BANCO';
      setPagoForm(prev => ({ ...prev, medio_pago: valorFinal }));
    } else {
      setPagoForm(prev => ({ ...prev, medio_pago: val }));
    }
  };

  const alEscribirBancoOtro = (txt) => {
    setBancoPersonalizado(txt);
    setPagoForm(prev => ({ ...prev, medio_pago: txt.toUpperCase().trim() || 'OTRO BANCO' }));
  };

  const obtenerMedioPagoLimpio = (abono) => {
    if (abono.medio_pago && abono.medio_pago.trim() !== '') {
      return abono.medio_pago.toUpperCase();
    }
    const ref = String(abono.referencia || abono.nota || '').toUpperCase();
    if (ref.includes('BRE-B') || ref.includes('BREB')) return 'BRE-B';
    if (ref.includes('BANCOLOMBIA') || ref.includes('TRANS')) return 'TRANSFERENCIA';
    if (ref.includes('BOGOTA') || ref.includes('BCO')) return 'BCO BOGOTA';
    if (ref.includes('NEQUI') || ref.includes('DAVIPLATA')) return 'NEQUI / DAVIPLATA';
    if (ref.includes('CHEQUE')) return 'CHEQUE';
    return 'EFECTIVO';
  };

  const abrirModalNuevo = () => {
    setPagoForm(prev => ({
      ...prev,
      id_editando: null,
      monto: '',
      referencia: '',
      fecha_pago: new Date().toISOString().split('T')[0],
      cliente_id: filtroClienteId || '',
      despacho_id: filtroDespachoId || ''
    }));
    setModalAbierto(true);
  };

  const abrirModalEditar = (abono) => {
    prepararEdicionPago(abono);
    setModalAbierto(true);
  };

  const handleGuardarAbonoDirecto = async (e) => {
    e.preventDefault();

    if (!pagoForm.cliente_id || !pagoForm.despacho_id || !pagoForm.monto) {
      if (mostrarAlerta) mostrarAlerta("Complete los campos obligatorios del abono", "error");
      return;
    }

    let medioFinal = 'EFECTIVO';
    if (modoMedio === 'OTRO_MANUAL') {
      medioFinal = bancoPersonalizado.toUpperCase().trim() || 'OTRO BANCO';
    } else {
      medioFinal = modoMedio.toUpperCase();
    }

    const montoNumerico = parseFloat(String(pagoForm.monto).replace(/\D/g, "")) || 0;
    if (montoNumerico <= 0) {
      if (mostrarAlerta) mostrarAlerta("El valor del abono debe ser mayor a cero", "error");
      return;
    }

    const payload = {
      cliente_id: pagoForm.cliente_id,
      despacho_id: pagoForm.despacho_id,
      fecha_pago: pagoForm.fecha_pago,
      monto: montoNumerico,
      medio_pago: medioFinal,
      referencia: pagoForm.referencia ? pagoForm.referencia.toUpperCase().trim() : null
    };

    try {
      if (pagoForm.id_editando) {
        const { error } = await supabase.from('pagos').update(payload).eq('id', pagoForm.id_editando);
        if (error) throw error;
        if (mostrarAlerta) mostrarAlerta(`Abono actualizado correctamente como [${medioFinal}]`, "exito");
      } else {
        const { error } = await supabase.from('pagos').insert([payload]);
        if (error) throw error;
        if (mostrarAlerta) mostrarAlerta(`Abono registrado con éxito como [${medioFinal}]`, "exito");
      }

      setFiltroClienteId(pagoForm.cliente_id);
      setFiltroDespachoId(pagoForm.despacho_id);

      setPagoForm(prev => ({ ...prev, monto: '', referencia: '', id_editando: null }));
      setModalAbierto(false); 

      if (cargarTodo) await cargarTodo();

    } catch (err) {
      console.error("Error guardando abono:", err);
      if (mostrarAlerta) mostrarAlerta("Error en base de datos: " + err.message, "error");
    }
  };

  // --- 📊 EXPORTAR PAGOS A EXCEL ---
  const exportarPagosAExcel = async () => {
    if (!datosPagos || datosPagos.length === 0) {
      if (typeof mostrarAlerta === "function") mostrarAlerta("No hay datos de pagos para exportar", "error");
      return;
    }

    try {
      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet('Control de Pagos');

      worksheet.columns = [
        { header: 'FECHA REMISIÓN', key: 'fecha_despacho', width: 15 },
        { header: 'INVERNADERO', key: 'invernadero', width: 18 },
        { header: 'CLIENTE', key: 'cliente', width: 25 },
        { header: 'NIT / CC', key: 'nit', width: 16 },
        { header: 'N° DE REMISIÓN', key: 'remision', width: 16 },
        { header: 'VALOR INICIAL', key: 'valor_inicial', width: 18 },
        { header: 'FECHA ABONO', key: 'fecha_abono', width: 15 },
        { header: 'MEDIO DE PAGO', key: 'medio_pago', width: 22 },
        { header: 'VALOR ABONO', key: 'valor_abono', width: 18 },
        { header: 'SALDO PENDIENTE', key: 'saldo', width: 18 },
        { header: 'N° COMPROBANTE / NOTA', key: 'nota', width: 30 }
      ];

      datosPagos.forEach((p) => {
        if (!p) return;
        const despachoCoincidente = datosDespachos?.find(d => d.id?.toString() === p.despacho_id?.toString());
        const fechaRemisionReal = despachoCoincidente?.fecha_venta || despachoCoincidente?.fecha || p.fecha_despacho || 'S/F';
        const invernaderoNom = despachoCoincidente?.invernaderos?.nombre || despachoCoincidente?.nombre_invernadero || 'GENERAL';
        const clienteNom = despachoCoincidente?.clientes?.nombre_completo || p.clientes?.nombre_completo || p.nombre_cliente || 'PARTICULAR';
        const clienteNit = despachoCoincidente?.clientes?.nit_cc || p.clientes?.nit_cc || p.nit_cc || 'N/A';
        const numeroRemision = despachoCoincidente?.numero_remision || p.numero_remision || 'S/N';
        const valorInicial = parseFloat(despachoCoincidente?.total_venta || p.valor_inicial || 0);
        const valorAbono = parseFloat(p.monto || p.valor_abono || 0);
        const saldoCalculado = valorInicial > 0 ? (valorInicial - valorAbono) : 0;

        worksheet.addRow({
          fecha_despacho: fechaRemisionReal ? String(fechaRemisionReal).split('T')[0] : '',
          invernadero: String(invernaderoNom).toUpperCase(),
          cliente: String(clienteNom).toUpperCase(),
          nit: clienteNit,
          remision: numeroRemision,
          valor_inicial: valorInicial,
          fecha_abono: p.fecha_pago ? String(p.fecha_pago).split('T')[0] : '',
          medio_pago: obtenerMedioPagoLimpio(p),
          valor_abono: valorAbono,
          saldo: saldoCalculado,
          nota: String(p.referencia || p.nota || '').toUpperCase()
        });
      });

      const ultFila = worksheet.rowCount;
      const filaTotales = worksheet.addRow({
        remision: 'TOTALES:',
        valor_inicial: { formula: `=SUM(F2:F${ultFila})` },
        valor_abono: { formula: `=SUM(I2:I${ultFila})` },
        saldo: { formula: `=SUM(J2:J${ultFila})` }
      });

      const headerRow = worksheet.getRow(1);
      headerRow.height = 24;
      headerRow.eachCell((cell) => {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF15803D' } };
        cell.font = { name: 'Arial', size: 10, bold: true, color: { argb: 'FFFFFFFF' } };
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
      });

      worksheet.eachRow((row, rowNumber) => {
        if (rowNumber === 1 || rowNumber === ultFila + 1) return;
        row.height = 20;
        const esCebra = rowNumber % 2 === 0;
        row.eachCell((cell, colNumber) => {
          cell.font = { name: 'Arial', size: 9 };
          if (esCebra) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE2EFDA' } };
          if ([1, 4, 5, 7, 8].includes(colNumber)) cell.alignment = { vertical: 'middle', horizontal: 'center' };
          else if ([6, 9, 10].includes(colNumber)) {
            cell.alignment = { vertical: 'middle', horizontal: 'right' };
            cell.numFmt = '"$"#,##0';
          } else cell.alignment = { vertical: 'middle', horizontal: 'left' };
        });
      });

      filaTotales.height = 22;
      filaTotales.eachCell((cell, colN) => {
        cell.font = { name: 'Arial', size: 10, bold: true, color: { argb: 'FF15803D' } };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD6EEFC' } };
        cell.border = { top: { style: 'thin' }, bottom: { style: 'double' } };
        if ([6, 9, 10].includes(colN)) {
          cell.alignment = { vertical: 'middle', horizontal: 'right' };
          cell.numFmt = '"$"#,##0';
        }
        if (colN === 5) cell.alignment = { vertical: 'middle', horizontal: 'right' };
      });

      worksheet.autoFilter = { from: { row: 1, column: 1 }, to: { row: ultFila, column: worksheet.columnCount } };
      const buffer = await workbook.xlsx.writeBuffer();
      const fechaHoy = new Date().toISOString().split('T')[0];
      const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      saveAs(blob, `REPORTE_PAGOS_CARTERA_${fechaHoy}.xlsx`);
      
      if (typeof mostrarAlerta === "function") mostrarAlerta("Reporte de cartera generado con éxito", "exito");
    } catch (error) {
      console.error("Error al exportar Excel de Pagos:", error);
    }
  };  

  // --- IMPRIMIR RECIBO CARTERA PDF ---
  const imprimirReciboCarteraPDF = async (remisionData) => {
    try {
      const remisionActiva = remisionData || remisionSeleccionada || {};
      if (!remisionActiva || Object.keys(remisionActiva).length === 0) {
        if (typeof mostrarAlerta === "function") mostrarAlerta("Por favor, seleccione una remisión antes de imprimir", "error");
        return;
      }

      const idVenta = remisionActiva.id;
      const nRemision = remisionActiva.numero_remision || 'S/N';
      const clienteNom = remisionActiva.clientes?.nombre_completo || remisionActiva.nombre_cliente || 'CLIENTE';
      const clienteNit = remisionActiva.clientes?.nit_cc || remisionActiva.nit_cc || 'N/A';
      const invernaderoNom = remisionActiva.invernaderos?.nombre || remisionActiva.nombre_invernadero || 'GENERAL';
      const fechaDespacho = remisionActiva.fecha_venta || remisionActiva.fecha || '';
      const notaRef = remisionActiva.referencia || remisionActiva.nota || 'SIN OBSERVACIONES';
      const valorTotalVenta = parseFloat(remisionActiva.total_venta || remisionActiva.total || 0);

      let bodyProductos = [];
      const { data: productosDB, error: errorProd } = await supabase.from('detalle_ventas').select('*').eq('venta_id', idVenta);

      if (!errorProd && productosDB && productosDB.length > 0) {
        bodyProductos = productosDB.map(item => [
          String(item.descripcion || 'PRODUCTO').toUpperCase(),
          `${item.cantidad || 0} ${item.escala || 'Unidad'}`,
          `$${parseFloat(item.subtotal || 0).toLocaleString('es-CO')}`
        ]);
      } else {
        const fallbackItems = remisionActiva.detalle_ventas || remisionActiva.items || [];
        if (fallbackItems.length > 0) {
          bodyProductos = fallbackItems.map(item => [
            String(item.descripcion || 'PRODUCTO').toUpperCase(),
            `${item.amount || item.cantidad || 0} ${item.escala || 'Unidad'}`,
            `$${parseFloat(item.subtotal || (item.cantidad * item.precio) || 0).toLocaleString('es-CO')}`
          ]);
        } else {
          bodyProductos = [[`PRODUCTOS DE LA REMISIÓN N° ${nRemision}`, "1 Global", `$${valorTotalVenta.toLocaleString('es-CO')}`]];
        }
      }

      let bodyAbonos = [];
      let totalAbonadoAcumulado = 0;
      const { data: pagosDB, error: errorPagos } = await supabase.from('pagos').select('*').eq('despacho_id', idVenta).order('fecha_pago', { ascending: true });

      if (!errorPagos && pagosDB && pagosDB.length > 0) {
        pagosDB.forEach((abono, index) => {
          const monto = parseFloat(abono.monto || 0);
          totalAbonadoAcumulado += monto;
          const medio = obtenerMedioPagoLimpio(abono);
          const ref = String(abono.referencia || abono.nota || 'ABONO REGISTRADO').toUpperCase();
          bodyAbonos.push([`${index + 1}`, abono.fecha_pago || 'S/F', `${medio} - ${ref}`, `$${monto.toLocaleString('es-CO')}`]);
        });
      } else {
        const pagosFallback = remisionActiva.pagos || remisionActiva.abonos || [];
        if (pagosFallback.length > 0) {
          pagosFallback.forEach((abono, index) => {
            const monto = parseFloat(abono.monto || 0);
            totalAbonadoAcumulado += monto;
            bodyAbonos.push([`${index + 1}`, abono.fecha_pago || 'S/F', String(abono.referencia || 'ABONO').toUpperCase(), `$${monto.toLocaleString('es-CO')}`]);
          });
        } else {
          bodyAbonos = [["-", "-", "SIN ABONOS REGISTRADOS A LA FECHA", "$0"]];
        }
      }

      const saldoNetoPendiente = valorTotalVenta - totalAbonadoAcumulado;
      const doc = new jsPDF({ orientation: 'portrait', unit: 'mm', format: [105, 148] });

      doc.setDrawColor(112, 173, 71); doc.setLineWidth(0.8); doc.rect(4, 4, 97, 140);
      try { doc.addImage('/Logopapel.png', 'PNG', 42.5, 6, 20, 20); } catch (e) {}

      doc.setFont("helvetica", "bold"); doc.setFontSize(8); doc.setTextColor(60, 60, 60); doc.text(`REMISIÓN N°: ${nRemision}`, 6, 11);
      doc.setFont("helvetica", "bold"); doc.setFontSize(11); doc.setTextColor(40, 80, 40); doc.text("ESTADO DE CUENTA DE CARTERA", 52.5, 29, { align: "center" });

      const yBase = 33; const altoFila = 5; const yOffset = 3.5;
      doc.setFillColor(242, 242, 242); doc.rect(6, yBase, 93, altoFila, 'F'); doc.rect(6, yBase + (altoFila * 2), 93, altoFila, 'F');
      doc.setDrawColor(210, 210, 210); doc.setLineWidth(0.2); doc.rect(6, yBase, 93, altoFila * 3);
      doc.line(6, yBase + altoFila, 99, yBase + altoFila); doc.line(6, yBase + (altoFila * 2), 99, yBase + (altoFila * 2)); doc.line(52, yBase, 52, yBase + altoFila);

      doc.setFont("helvetica", "bold"); doc.setFontSize(7.5); doc.setTextColor(0); doc.text("FECHA DESPACHO:", 8, yBase + yOffset);
      doc.setFont("helvetica", "normal"); doc.text(`${fechaDespacho}`, 34, yBase + yOffset);
      doc.setFont("helvetica", "bold"); doc.text("INVERNADERO:", 54, yBase + yOffset);
      doc.setFont("helvetica", "normal"); doc.text(`${invernaderoNom.toUpperCase()}`, 75, yBase + yOffset);
      doc.setFont("helvetica", "bold"); doc.text("CLIENTE:", 8, yBase + altoFila + yOffset);
      doc.setFont("helvetica", "normal"); doc.text(`${clienteNom.toUpperCase()}`, 22, yBase + altoFila + yOffset);
      doc.setFont("helvetica", "bold"); doc.text("NIT / CC:", 8, yBase + (altoFila * 2) + yOffset);
      doc.setFont("helvetica", "normal"); doc.text(`${clienteNit}`, 22, yBase + (altoFila * 2) + yOffset);

      autoTable(doc, {
        startY: yBase + (altoFila * 3) + 3, margin: { left: 6, right: 6 },
        head: [["PRODUCTO DESPACHADO", "CANTIDAD", "SUBTOTAL"]], body: bodyProductos, theme: 'grid',
        styles: { font: 'helvetica', fontSize: 7, cellPadding: 1.5, lineWidth: 0.1, lineColor: [210, 210, 210] },
        headStyles: { fillColor: [112, 173, 71], textColor: [255, 255, 255], halign: 'center', fontStyle: 'bold' },
        columnStyles: { 0: { cellWidth: 45, halign: 'left' }, 1: { cellWidth: 25, halign: 'center' }, 2: { cellWidth: 23, halign: 'right' } }
      });

      const yTotalVenta = doc.lastAutoTable.finalY + 4;
      doc.setFont("helvetica", "bold"); doc.setFontSize(8); doc.text("VALOR TOTAL VENTA:", 53, yTotalVenta);
      doc.text(`$${valorTotalVenta.toLocaleString('es-CO')}`, 99, yTotalVenta, { align: 'right' });

      autoTable(doc, {
        startY: yTotalVenta + 2, margin: { left: 6, right: 6 },
        head: [["N°", "FECHA PAGO", "MEDIO / REFERENCIA", "VALOR ABONO"]], body: bodyAbonos, theme: 'grid',
        styles: { font: 'helvetica', fontSize: 7, cellPadding: 1.5, lineWidth: 0.1, lineColor: [210, 210, 210] },
        headStyles: { fillColor: [40, 80, 40], textColor: [255, 255, 255], halign: 'center', fontStyle: 'bold' },
        columnStyles: { 0: { cellWidth: 8, halign: 'center' }, 1: { cellWidth: 20, halign: 'center' }, 2: { cellWidth: 42, halign: 'left' }, 3: { cellWidth: 23, halign: 'right' } }
      });

      const yBalance = doc.lastAutoTable.finalY + 4;
      doc.setFillColor(245, 245, 245); doc.rect(48, yBalance, 51, 11, 'F');
      doc.setDrawColor(200, 200, 200); doc.rect(48, yBalance, 51, 11); doc.line(48, yBalance + 5.5, 99, yBalance + 5.5);

      doc.setFont("helvetica", "bold"); doc.setFontSize(7.5); doc.setTextColor(0); doc.text("TOTAL ABONADO:", 50, yBalance + 4);
      doc.text(`$${totalAbonadoAcumulado.toLocaleString('es-CO')}`, 97, yBalance + 4, { align: 'right' });
      doc.setFont("helvetica", "bold"); doc.text("SALDO PENDIENTE:", 50, yBalance + 9.5);
      doc.setFont("helvetica", "bold"); doc.setTextColor(180, 0, 0); doc.text(`$${saldoNetoPendiente.toLocaleString('es-CO')}`, 97, yBalance + 9.5, { align: 'right' });

      const yNotas = yBalance + 14; doc.setDrawColor(210, 210, 210); doc.rect(6, yNotas, 93, 7);
      doc.setFont("helvetica", "bold"); doc.setFontSize(6.5); doc.setTextColor(0); doc.text("NOTA:", 7, yNotas + 4.5);
      doc.setFont("helvetica", "normal"); doc.text(`${String(notaRef).toUpperCase()}`, 15, yNotas + 4.5, { maxWidth: 82 });

      doc.save(`ESTADO_CUENTA_REM_${nRemision}.pdf`);
    } catch (err) {
      console.error(err);
    }
  };

  const remisionSeleccionada = datosDespachos?.find(r => r.id?.toString() === filtroDespachoId?.toString());
  const historialAbonos = datosPagos
    ?.filter(p => p.despacho_id?.toString() === filtroDespachoId?.toString())
    .sort((a, b) => new Date(a.fecha_pago) - new Date(b.fecha_pago));

  const totalAbonado = historialAbonos?.reduce((acc, p) => acc + (parseFloat(p.monto) || 0), 0) || 0;
  const saldoActual = remisionSeleccionada ? (parseFloat(remisionSeleccionada.total_venta) - totalAbonado) : 0;
  const remisionesDelCliente = datosDespachos?.filter(d => d.cliente_id?.toString() === filtroClienteId?.toString()) || [];

  return (
    <div className="space-y-6 pb-20 text-slate-800 dark:text-slate-200 font-sans transition-colors duration-300">
      
      {/* 🚀 CABECERA PRINCIPAL MODERNA */}
      <div className="bg-white dark:bg-slate-800 p-6 rounded-3xl shadow-xl border border-slate-200 dark:border-slate-700 flex flex-col xl:flex-row justify-between items-center gap-4 transition-colors duration-300">
        <div className="space-y-1 w-full xl:w-auto text-center xl:text-left">
          <div className="flex items-center justify-center xl:justify-start gap-3">
            <span className="text-2xl p-2 bg-blue-700/10 dark:bg-blue-500/20 rounded-xl text-blue-700 dark:text-blue-400">💳</span>
            <h2 className="text-xl font-black tracking-tight text-slate-900 dark:text-white uppercase">Pagos y Cartera</h2>
          </div>
          <p className="text-xs text-slate-500 dark:text-slate-400 font-medium">Registro de abonos, cuentas por cobrar y estados de cuenta.</p>
        </div>
        
        {/* BOTONERA DE ACCIONES RÁPIDAS */}
        <div className="flex items-center gap-3 flex-wrap justify-center xl:justify-end w-full xl:w-auto">
          <button onClick={exportarPagosAExcel} className="flex-1 md:flex-none px-4 py-2.5 bg-emerald-600 hover:bg-emerald-700 text-white font-black text-xs uppercase tracking-wider rounded-xl shadow-md transition-all flex items-center justify-center gap-1.5 cursor-pointer">
            <span>📊</span> Exportar Excel
          </button>
          <button onClick={abrirModalNuevo} disabled={!filtroDespachoId} className={`flex-1 md:flex-none px-5 py-2.5 font-black text-xs uppercase tracking-wider rounded-xl shadow-md transition-all flex items-center justify-center gap-1.5 ${filtroDespachoId ? 'bg-blue-700 hover:bg-blue-800 text-white cursor-pointer' : 'bg-slate-200 dark:bg-slate-700 text-slate-400 dark:text-slate-500 cursor-not-allowed'}`}>
            <span>💰</span> Registrar Abono
          </button>
        </div>
      </div>

      {/* 🔍 BARRA DE FILTROS SUPERIOR */}
      <div className="bg-white dark:bg-slate-800 p-5 rounded-2xl shadow-md border border-slate-200 dark:border-slate-700 grid grid-cols-1 md:grid-cols-2 gap-4 transition-colors duration-300">
        <div>
          <label className="text-[10px] font-black text-slate-500 dark:text-slate-400 uppercase tracking-wider mb-1 block">👤 Seleccionar Cliente</label>
          <select 
            className="w-full bg-slate-50 dark:bg-slate-700/50 border border-slate-200 dark:border-slate-600 p-2.5 rounded-xl font-bold text-xs uppercase outline-none focus:border-blue-500 dark:text-white"
            value={filtroClienteId}
            onChange={(e) => {
              setFiltroClienteId(e.target.value);
              setFiltroDespachoId(''); 
            }} 
          >
            <option value="">SELECCIONE CLIENTE PARA VER CARTERA...</option>
            {listaClientes?.filter(c => c.activo !== false).map(c => (
              <option key={c.id} value={c.id}>{c.nombre_completo?.toUpperCase()}</option>
            ))}
          </select>
        </div>

        <div>
          <label className="text-[10px] font-black text-slate-500 dark:text-slate-400 uppercase tracking-wider mb-1 block">📄 Seleccionar N° de Remisión / Venta</label>
          <select 
            className="w-full bg-slate-50 dark:bg-slate-700/50 border border-slate-200 dark:border-slate-600 p-2.5 rounded-xl font-bold text-xs uppercase outline-none focus:border-blue-500 disabled:opacity-50 disabled:cursor-not-allowed dark:text-white"
            value={filtroDespachoId}
            onChange={(e) => setFiltroDespachoId(e.target.value)}
            disabled={!filtroClienteId} 
          >
            <option value="">{filtroClienteId ? 'SELECCIONE LA REMISIÓN...' : 'ESPERANDO CLIENTE...'}</option>
            {remisionesDelCliente.map(r => (
              <option key={r.id} value={r.id}>N° {r.numero_remision} - Total: {formatoPesos(r.total_venta)}</option>
            ))}
          </select>
        </div>
      </div>

      {/* 📊 FICHA DE ESTADO DE CUENTA (ANCHO COMPLETO) */}
      <div className="w-full">
        {remisionSeleccionada ? (
          <div className="bg-white dark:bg-slate-800 rounded-3xl shadow-xl overflow-hidden border border-gray-200 dark:border-slate-700 transition-colors duration-300">
            
            <div className="bg-slate-800 dark:bg-slate-900 p-5 text-white flex flex-col md:flex-row justify-between items-center gap-4">
              <div className="flex items-center gap-3">
                <span className="text-2xl">📄</span>
                <div>
                  <h3 className="font-black uppercase text-base tracking-widest italic">Estado de Cuenta: Remisión N° {remisionSeleccionada.numero_remision}</h3>
                  <p className="text-[10px] text-slate-400 font-bold uppercase mt-0.5">Cliente: {remisionSeleccionada.clientes?.nombre_completo}</p>
                </div>
              </div>
              
              <div className="flex items-center gap-2">
                {saldoActual <= 0 ? (
                  <span className="bg-emerald-500/20 border border-emerald-500 text-emerald-400 px-3 py-1 rounded-lg text-[10px] font-black uppercase">🟢 PAGADA EN SU TOTALIDAD</span>
                ) : totalAbonado > 0 ? (
                  <span className="bg-amber-500/20 border border-amber-500 text-amber-400 px-3 py-1 rounded-lg text-[10px] font-black uppercase">🟡 PARCIALMENTE ABONADA</span>
                ) : (
                  <span className="bg-rose-500/20 border border-rose-500 text-rose-400 px-3 py-1 rounded-lg text-[10px] font-black uppercase">🔴 PENDIENTE DE PAGO</span>
                )}
              </div>
            </div>

            <div className="p-6 space-y-6">
              {/* BLOQUE INFORMATIVO DE LA VENTA */}
              <div className="grid grid-cols-1 lg:grid-cols-2 gap-6 border-b border-slate-100 dark:border-slate-700 pb-6">
                <div className="space-y-1">
                  <p className="text-[10px] font-black text-gray-400 uppercase tracking-widest">Información del Despacho</p>
                  <p className="font-black text-xl text-slate-900 dark:text-white uppercase leading-tight">{remisionSeleccionada.clientes?.nombre_completo}</p>
                  <div className="flex gap-4 pt-1">
                    <p className="text-[11px] font-bold text-slate-500 dark:text-slate-400 uppercase">📅 Carga: <span className="text-slate-800 dark:text-slate-200">{remisionSeleccionada.fecha_venta}</span></p>
                    <p className="text-[11px] font-bold text-slate-500 dark:text-slate-400 uppercase">🌱 Invernadero: <span className="text-slate-800 dark:text-slate-200">{remisionSeleccionada.invernaderos?.nombre || 'GENERAL'}</span></p>
                  </div>
                </div>

                <div className="bg-blue-50/50 dark:bg-slate-700/50 p-4 rounded-2xl border border-blue-100 dark:border-slate-600 shadow-sm">
                  <p className="text-[10px] font-black text-blue-700 dark:text-sky-400 uppercase mb-3 tracking-wider flex items-center gap-1.5 border-b border-blue-200/50 dark:border-slate-600 pb-1.5">
                    <span>📦</span> Contenido Despachado
                  </p>
                  <div className="space-y-2">
                    {remisionSeleccionada.detalle_ventas?.map((item, i) => (
                      <div key={i} className="flex justify-between items-center pb-1">
                        <p className="text-xs font-bold text-slate-700 dark:text-slate-300 uppercase">{item.descripcion}</p>
                        <p className="text-[10px] font-black text-blue-800 dark:text-sky-300 bg-blue-100/80 dark:bg-sky-950/80 px-2 py-0.5 rounded-md border border-blue-200 dark:border-sky-800">
                          {item.amount || item.cantidad} {item.escala}
                        </p>
                      </div>
                    ))}
                    <div className="pt-3 border-t border-dashed border-blue-200 dark:border-slate-600 flex justify-between items-center">
                      <p className="text-[10px] font-black text-slate-500 dark:text-slate-400 uppercase tracking-widest">Valor Total Venta:</p>
                      <p className="text-lg font-black text-slate-900 dark:text-white">{formatoPesos(remisionSeleccionada.total_venta)}</p>
                    </div>
                  </div>
                </div>
              </div>

              {/* TABLA DE ABONOS RECIBIDOS A ANCHO COMPLETO */}
              <div className="space-y-3">
                <p className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Historial de Abonos Recibidos</p>
                
                {historialAbonos?.length > 0 ? (
                  <div className="overflow-x-auto rounded-2xl border border-slate-200 dark:border-slate-700 shadow-sm">
                    <table className="w-full text-left text-xs border-collapse">
                      <thead>
                        <tr className="bg-slate-100 dark:bg-slate-700/80 text-slate-600 dark:text-slate-300 uppercase font-black text-[10px] tracking-wider border-b border-slate-200 dark:border-slate-700">
                          <th className="p-4 text-center w-12">N°</th>
                          <th className="p-4">Fecha Abono</th>
                          <th className="p-4 text-center">Forma / Medio de Pago</th>
                          <th className="p-4">N° Ref / Comprobante</th>
                          <th className="p-4 text-right">Monto Abono</th>
                          <th className="p-4 text-center">Acciones</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-slate-100 dark:divide-slate-700/60 font-bold text-slate-700 dark:text-slate-300">
                        {historialAbonos.map((abono, idx) => {
                          const medioLimpio = obtenerMedioPagoLimpio(abono);

                          return (
                            <tr key={abono.id} className={`${idx % 2 === 0 ? 'bg-white dark:bg-slate-800' : 'bg-slate-50/50 dark:bg-slate-800/60'} hover:bg-sky-50/50 dark:hover:bg-slate-700/50 transition-colors border-l-4 border-blue-500`}>
                              <td className="p-4 text-center">
                                <span className="bg-blue-100 dark:bg-blue-950 text-blue-800 dark:text-blue-300 w-6 h-6 rounded-full inline-flex items-center justify-center font-black text-[10px] shadow-sm">
                                  {idx + 1}
                                </span>
                              </td>

                              <td className="p-4 font-black text-slate-900 dark:text-white whitespace-nowrap">
                                {abono.fecha_pago}
                              </td>

                              <td className="p-4 text-center">
                                <span className="bg-emerald-50 dark:bg-emerald-950/60 text-emerald-700 dark:text-emerald-400 border border-emerald-200 dark:border-emerald-800 px-2.5 py-1 rounded-md text-[9px] font-black uppercase tracking-tight inline-block shadow-sm">
                                  💳 {medioLimpio}
                                </span>
                              </td>

                              <td className="p-4 uppercase text-slate-500 dark:text-slate-400 font-bold text-[10px]">
                                {abono.referencia || abono.nota ? (
                                  <span className="bg-slate-100 dark:bg-slate-700 border border-slate-200 dark:border-slate-600 px-2.5 py-1 rounded-md text-slate-700 dark:text-slate-200 font-black shadow-sm">
                                    {abono.referencia || abono.nota}
                                  </span>
                                ) : (
                                  <span className="text-gray-300 dark:text-slate-600 italic font-medium">Sin referencia</span>
                                )}
                              </td>

                              <td className="p-4 text-right font-black text-blue-700 dark:text-sky-400 text-sm whitespace-nowrap">
                                +{formatoPesos(abono.monto)}
                              </td>

                              <td className="p-4 text-center">
                                <div className="flex gap-1.5 justify-center">
                                  <button 
                                    type="button"
                                    onClick={() => abrirModalEditar(abono)}
                                    className="p-1.5 bg-slate-700 dark:bg-slate-600 text-white rounded-lg shadow-sm border border-slate-800 hover:bg-slate-900 transition-colors text-[10px] font-black flex items-center gap-1 cursor-pointer"
                                    title="Editar Abono"
                                  >
                                    <span>✏️</span> EDITAR
                                  </button>
                                  <button 
                                    type="button"
                                    onClick={() => eliminarPago(abono.id)}
                                    className="p-1.5 bg-red-100 dark:bg-red-950/60 text-red-700 dark:text-red-400 rounded-lg hover:bg-red-700 hover:text-white border border-red-200 dark:border-red-900 transition-colors cursor-pointer"
                                    title="Eliminar Abono"
                                  >
                                    🗑️
                                  </button>
                                </div>
                              </td>
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>
                  </div>
                ) : (
                  <div className="text-center py-8 bg-slate-50 dark:bg-slate-800/40 rounded-2xl border border-dashed border-slate-300 dark:border-slate-700">
                    <span className="text-3xl grayscale opacity-50 block mb-2">💸</span>
                    <p className="text-[11px] font-black text-slate-400 dark:text-slate-500 uppercase tracking-widest">Sin abonos registrados a esta remisión</p>
                  </div>
                )}
              </div>

              {/* BLOQUE INFERIOR DE TOTALES Y PDF */}
              <div className="pt-4 flex flex-col md:flex-row gap-4 items-stretch md:items-end justify-between">
                
                <div className="bg-slate-900 dark:bg-slate-950 p-5 rounded-2xl flex-1 flex justify-between items-center text-white shadow-xl border border-slate-800">
                  <div>
                    <p className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Total Recaudado</p>
                    <p className="font-black text-xl text-blue-400 mt-0.5">{formatoPesos(totalAbonado)}</p>
                  </div>
                  <div className="text-right">
                    <p className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Saldo Neto Pendiente</p>
                    <p className={`font-black text-3xl mt-0.5 tracking-tight ${saldoActual <= 0 ? 'text-green-400' : 'text-red-400'}`}>
                      {formatoPesos(saldoActual)}
                    </p>
                  </div>
                </div>

                <button
                  type="button"
                  onClick={() => imprimirReciboCarteraPDF(remisionSeleccionada)}
                  className="px-6 py-4 h-full bg-rose-700 hover:bg-rose-800 text-white font-black rounded-2xl shadow-xl transition-colors flex items-center justify-center gap-2 text-xs uppercase tracking-wider border border-rose-600 cursor-pointer"
                >
                  <span className="text-lg">🖨️</span> Imprimir PDF Remisión
                </button>

              </div>
            </div>
          </div>
        ) : (
          <div className="bg-slate-50 dark:bg-slate-800/40 border-2 border-dashed border-slate-200 dark:border-slate-700 p-16 rounded-3xl text-center flex flex-col items-center justify-center h-full min-h-[400px]">
            <span className="text-5xl grayscale opacity-40 mb-4">📂</span>
            <p className="text-slate-400 dark:text-slate-500 font-black uppercase text-sm tracking-widest">Seleccione un cliente y una remisión</p>
            <p className="text-slate-400 dark:text-slate-600 text-xs mt-2 font-bold">Para cargar el desglose y su estado de cuenta de cartera</p>
          </div>
        )}
      </div>

      {/* 🧊 MODAL FLOTANTE PARA REGISTRAR O EDITAR ABONOS */}
      {modalAbierto && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 backdrop-blur-sm p-4 overflow-y-auto">
          <div className="bg-white dark:bg-slate-800 border border-slate-200 dark:border-slate-700 w-full max-w-lg rounded-3xl shadow-2xl overflow-hidden animate-in fade-in zoom-in-95 duration-200 my-8 text-slate-800 dark:text-slate-200">
            
            <div className="p-5 bg-slate-900 text-white border-b border-slate-800 flex justify-between items-center">
              <div className="flex items-center gap-3">
                <span className="text-xl">💳</span>
                <div>
                  <h3 className="text-base font-black uppercase tracking-tight">
                    {pagoForm.id_editando ? 'Editar Abono Existente' : 'Registrar Nuevo Abono'}
                  </h3>
                  <p className="text-[10px] text-slate-400 font-medium">Asignación de pagos a la remisión N° {remisionSeleccionada?.numero_remision}</p>
                </div>
              </div>
              <button onClick={() => setModalAbierto(false)} className="text-slate-400 hover:text-white text-lg font-black px-3 py-1 rounded-xl bg-slate-800 cursor-pointer">✕</button>
            </div>

            <form onSubmit={handleGuardarAbonoDirecto} className="p-6 space-y-4">
              
              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                <div>
                  <label className="text-[10px] font-black text-slate-500 dark:text-slate-400 uppercase tracking-wider">Fecha Abono *</label>
                  <input 
                    type="date" 
                    className="w-full bg-slate-50 dark:bg-slate-900 border border-slate-200 dark:border-slate-700 text-slate-800 dark:text-white p-3 rounded-xl font-bold text-sm outline-none focus:border-blue-500 mt-1"
                    value={pagoForm.fecha_pago}
                    onChange={(e) => setPagoForm({...pagoForm, fecha_pago: e.target.value})} 
                    required 
                  />
                </div>

                <div>
                  <label className="text-[10px] font-black text-slate-500 dark:text-slate-400 uppercase tracking-wider">Medio / Forma de Pago *</label>
                  <select 
                    className="w-full bg-slate-50 dark:bg-slate-900 border border-slate-200 dark:border-slate-700 text-slate-800 dark:text-white p-3 rounded-xl font-bold text-xs outline-none focus:border-blue-500 uppercase mt-1"
                    value={modoMedio}
                    onChange={(e) => alCambiarMedioSelect(e.target.value)}
                    required
                  >
                    {listaMediosPredeterminados.map(m => (
                      <option key={m} value={m}>
                        {m === 'OTRO_MANUAL' ? '✏️ OTRO BANCO (Escribir...)' : m}
                      </option>
                    ))}
                  </select>
                </div>
              </div>

              {modoMedio === 'OTRO_MANUAL' && (
                <div className="bg-amber-50 dark:bg-amber-950/40 p-4 rounded-xl border border-amber-300 dark:border-amber-700 shadow-inner animate-in fade-in">
                  <label className="text-[9px] font-black text-amber-900 dark:text-amber-400 uppercase tracking-wider block mb-1">Nombre del Banco / Entidad *</label>
                  <input 
                    type="text" 
                    className="w-full p-2.5 bg-white dark:bg-slate-900 border border-amber-300 dark:border-amber-700 rounded-lg font-black text-xs uppercase outline-none focus:border-amber-600 text-amber-900 dark:text-amber-200"
                    placeholder="Ej: BCO BOGOTA / DAVIVIENDA"
                    value={bancoPersonalizado}
                    onChange={(e) => alEscribirBancoOtro(e.target.value)}
                    required
                  />
                </div>
              )}

              <div className="bg-blue-50/50 dark:bg-slate-700/40 p-4 rounded-2xl border border-blue-200 dark:border-slate-600 shadow-inner space-y-3">
                <div>
                  <label className="text-[10px] font-black text-blue-800 dark:text-sky-400 uppercase tracking-wider">Valor del Abono *</label>
                  <input 
                    type="text" 
                    className="w-full p-3 bg-white dark:bg-slate-900 border-2 border-blue-200 dark:border-slate-600 rounded-xl font-black text-2xl text-blue-900 dark:text-white outline-none focus:border-blue-500 mt-1 placeholder:text-blue-200 dark:placeholder:text-slate-600"
                    value={formatearMascaraMoneda(pagoForm.monto)} 
                    onChange={(e) => setPagoForm({...pagoForm, monto: e.target.value.replace(/\D/g, "")})} 
                    placeholder="$ 0"
                    required 
                  />
                </div>
                
                {!pagoForm.id_editando && (
                  <div className="flex justify-between items-center pt-2 border-t border-blue-100 dark:border-slate-600">
                    <p className="text-[10px] font-black text-slate-500 dark:text-slate-400 uppercase tracking-widest">Saldo Pendiente Actual:</p>
                    <p className={`font-black text-sm ${saldoActual <= 0 ? 'text-emerald-600 dark:text-emerald-400' : 'text-rose-600 dark:text-rose-400'}`}>
                      {formatoPesos(saldoActual)}
                    </p>
                  </div>
                )}
              </div>

              <div>
                <label className="text-[10px] font-black text-slate-500 dark:text-slate-400 uppercase tracking-wider">N° Comprobante / Referencia</label>
                <input 
                  type="text"
                  className="w-full bg-slate-50 dark:bg-slate-900 border border-slate-200 dark:border-slate-700 text-slate-800 dark:text-white p-3 rounded-xl font-bold text-xs outline-none focus:border-blue-500 uppercase mt-1 placeholder-slate-400"
                  value={pagoForm.referencia || ''}
                  onChange={(e) => setPagoForm({...pagoForm, referencia: e.target.value})}
                  placeholder="Ej: N° Transacción 458921 / Cheque N° 102" 
                />
              </div>

              <div className="flex justify-end gap-3 pt-4 border-t border-slate-200 dark:border-slate-700">
                <button type="button" onClick={() => setModalAbierto(false)} className="px-5 py-2.5 bg-slate-200 dark:bg-slate-700 hover:bg-slate-300 dark:hover:bg-slate-600 text-slate-700 dark:text-slate-200 font-black text-xs uppercase tracking-wider rounded-xl transition-colors cursor-pointer">
                  Cancelar
                </button>
                <button type="submit" className={`px-6 py-2.5 text-white font-black text-xs uppercase tracking-wider rounded-xl shadow-lg transition-colors cursor-pointer ${pagoForm.id_editando ? 'bg-amber-600 hover:bg-amber-700' : 'bg-blue-700 hover:bg-blue-800'}`}>
                  {pagoForm.id_editando ? '💾 Actualizar Abono' : '💰 Registrar Abono'}
                </button>
              </div>
            </form>

          </div>
        </div>
      )}

    </div>
  );
}