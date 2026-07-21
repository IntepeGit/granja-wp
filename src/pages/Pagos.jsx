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

  // Manejador del cambio en el selector de Medio de Pago
  const alCambiarMedioSelect = (val) => {
    setModoMedio(val);
    if (val === 'OTRO_MANUAL') {
      const valorFinal = bancoPersonalizado.toUpperCase().trim() || 'OTRO BANCO';
      setPagoForm(prev => ({ ...prev, medio_pago: valorFinal }));
    } else {
      setPagoForm(prev => ({ ...prev, medio_pago: val }));
    }
  };

  // Manejador de la escritura manual del banco
  const alEscribirBancoOtro = (txt) => {
    setBancoPersonalizado(txt);
    setPagoForm(prev => ({ ...prev, medio_pago: txt.toUpperCase().trim() || 'OTRO BANCO' }));
  };

  // Detección inteligente de medio de pago para registros anteriores
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

  // ⚡ GUARDADO DIRECTO CONSERVA LA SELECCIÓN DE LA REMISIÓN
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
        const { error } = await supabase
          .from('pagos')
          .update(payload)
          .eq('id', pagoForm.id_editando);

        if (error) throw error;
        if (mostrarAlerta) mostrarAlerta(`Abono actualizado correctamente como [${medioFinal}]`, "exito");
      } else {
        const { error } = await supabase
          .from('pagos')
          .insert([payload]);

        if (error) throw error;
        if (mostrarAlerta) mostrarAlerta(`Abono registrado con éxito como [${medioFinal}]`, "exito");
      }

      // 🌟 MANTENEMOS EL CLIENTE Y EL DESPACHO PARA NO PERDER LA VISTA
      setPagoForm(prev => ({
        ...prev,
        monto: '',
        referencia: '',
        id_editando: null
      }));

      // Recargamos los datos para refrescar la lista a la derecha sin cerrar la ficha
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

  const remisionSeleccionada = datosDespachos?.find(r => r.id?.toString() === pagoForm.despacho_id?.toString());
  const historialAbonos = datosPagos
    ?.filter(p => p.despacho_id?.toString() === pagoForm.despacho_id?.toString())
    .sort((a, b) => new Date(a.fecha_pago) - new Date(b.fecha_pago));

  const totalAbonado = historialAbonos?.reduce((acc, p) => acc + (parseFloat(p.monto) || 0), 0) || 0;
  const saldoActual = remisionSeleccionada ? (parseFloat(remisionSeleccionada.total_venta) - totalAbonado) : 0;
  const remisionesDelCliente = datosDespachos?.filter(d => d.cliente_id?.toString() === pagoForm.cliente_id?.toString()) || [];

  return (
    <div className="space-y-6 pb-20 text-slate-800">
      <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
        
        {/* COLUMNA IZQUIERDA: FORMULARIO DE RECAUDO */}
        <div className="bg-white p-6 rounded-3xl shadow-xl border-t-8 border-blue-700 h-fit space-y-5">
          
          <div className="flex justify-between items-center border-b pb-3">
            <h3 className="font-black text-slate-800 uppercase text-xs italic">💳 Registro de Pagos / Abonos</h3>
          </div>

          <form onSubmit={handleGuardarAbonoDirecto} className="space-y-4">
            <div>
              <label className="text-[10px] font-bold text-gray-400 uppercase px-1 italic">Cliente *</label>
              <select 
                className="w-full border-2 p-2.5 rounded-xl font-bold text-xs bg-white outline-none focus:border-blue-500 uppercase"
                value={pagoForm.cliente_id}
                onChange={(e) => setPagoForm({...pagoForm, cliente_id: e.target.value, despacho_id: ''})} 
                required
              >
                <option value="">Seleccione Cliente...</option>
                {listaClientes?.filter(c => c.activo !== false).map(c => (
                  <option key={c.id} value={c.id}>{c.nombre_completo?.toUpperCase()}</option>
                ))}
              </select>
            </div>

            <div>
              <label className="text-[10px] font-bold text-gray-400 uppercase px-1 italic">N° Remisión / Venta *</label>
              <select 
                className="w-full border-2 p-2.5 rounded-xl font-bold text-xs bg-white outline-none focus:border-blue-500 uppercase"
                value={pagoForm.despacho_id}
                onChange={(e) => setPagoForm({...pagoForm, despacho_id: e.target.value, monto: ''})}
                disabled={!pagoForm.cliente_id} 
                required
              >
                <option value="">Seleccione Remisión...</option>
                {remisionesDelCliente.map(r => (
                  <option key={r.id} value={r.id}>N° {r.numero_remision} - Total: {formatoPesos(r.total_venta)}</option>
                ))}
              </select>
            </div>

            <div>
              <label className="text-[10px] font-bold text-gray-400 uppercase px-1 italic">Fecha Abono *</label>
              <input 
                type="date" 
                className="w-full border-2 p-2.5 rounded-xl font-bold text-sm outline-none focus:border-blue-500"
                value={pagoForm.fecha_pago}
                onChange={(e) => setPagoForm({...pagoForm, fecha_pago: e.target.value})} 
                required 
              />
            </div>

            {/* SELECTOR MEDIO DE PAGO */}
            <div>
              <label className="text-[10px] font-bold text-gray-400 uppercase px-1 italic">Medio / Forma de Pago *</label>
              <select 
                className="w-full border-2 p-2.5 rounded-xl font-bold text-xs bg-white outline-none focus:border-blue-500 uppercase"
                value={modoMedio}
                onChange={(e) => alCambiarMedioSelect(e.target.value)}
                required
              >
                {listaMediosPredeterminados.map(m => (
                  <option key={m} value={m}>
                    {m === 'OTRO_MANUAL' ? '✏️ OTRO BANCO / ENTIDAD (Escribir...)' : m}
                  </option>
                ))}
              </select>
            </div>

            {modoMedio === 'OTRO_MANUAL' && (
              <div className="bg-amber-50 p-2.5 rounded-xl border border-amber-300">
                <label className="text-[9px] font-black text-amber-900 uppercase px-1 italic block mb-1">Nombre del Banco / Entidad *</label>
                <input 
                  type="text" 
                  className="w-full p-2 border-2 bg-white rounded-lg font-black text-xs uppercase outline-none focus:border-amber-600 text-amber-900"
                  placeholder="Ej: BCO BOGOTA / DAVIVIENDA"
                  value={bancoPersonalizado}
                  onChange={(e) => alEscribirBancoOtro(e.target.value)}
                  required
                />
              </div>
            )}

            {remisionSeleccionada && (
              <div className="bg-blue-50/50 p-3 rounded-2xl border border-blue-100 space-y-3 shadow-inner">
                <div>
                  <label className="text-[9px] font-black text-blue-700 uppercase px-1 italic">Valor del Nuevo Abono *</label>
                  <input 
                    type="text" 
                    className="w-full p-2.5 border-2 bg-white rounded-xl font-black text-lg text-blue-900 border-blue-200 outline-none focus:border-blue-500"
                    value={formatearMascaraMoneda(pagoForm.monto)} 
                    onChange={(e) => setPagoForm({...pagoForm, monto: e.target.value.replace(/\D/g, "")})} 
                    required 
                  />
                </div>
                <div>
                  <label className="text-[9px] font-black text-slate-500 uppercase px-1 italic">Saldo Pendiente Actual</label>
                  <div className="w-full p-2 bg-white rounded-xl border-2 border-slate-200 flex items-center justify-center font-black text-sm">
                    <p className={saldoActual <= 0 ? 'text-green-600' : 'text-red-600'}>
                      {formatoPesos(saldoActual)}
                    </p>
                  </div>
                </div>
              </div>
            )}

            <div>
              <label className="text-[10px] font-bold text-gray-400 uppercase px-1 italic">N° Comprobante / Referencia</label>
              <input 
                className="w-full border-2 p-2.5 rounded-xl font-bold text-xs bg-white outline-none focus:border-blue-500 uppercase"
                value={pagoForm.referencia || ''}
                onChange={(e) => setPagoForm({...pagoForm, referencia: e.target.value})}
                placeholder="Ej: N° Transacción 458921 / Cheque N° 102" 
              />
            </div>

            <button 
              type="submit" 
              className={`w-full p-3.5 rounded-xl font-black uppercase tracking-wider text-xs transition-colors shadow-md cursor-pointer ${
                pagoForm.id_editando ? 'bg-amber-600 hover:bg-amber-700' : 'bg-green-700 hover:bg-green-800'
              } text-white`}
            >
              {pagoForm.id_editando ? '💾 Actualizar Abono' : '💰 Registrar Abono'}
            </button>
          </form>

          <div className="pt-2 border-t border-slate-100">
            <button
              onClick={exportarPagosAExcel}
              className="w-full py-3 bg-emerald-600 hover:bg-emerald-700 text-white font-black rounded-xl shadow-md transition-all flex items-center justify-center gap-2 text-[10px] uppercase tracking-wider border border-emerald-500 cursor-pointer"
            >
              📊 EXPORTAR A EXCEL TOTAL REGISTRO DE PAGOS
            </button>
          </div>
        </div>

        {/* COLUMNA DERECHA: GRILLA DISTRIBUIDA EN COLUMNAS */}
        <div className="lg:col-span-2 space-y-6">
          {remisionSeleccionada ? (
            <div className="bg-white rounded-3xl shadow-xl overflow-hidden border border-gray-200">
              
              <div className="bg-slate-800 p-4 text-white flex justify-between items-center flex-wrap gap-2">
                <div className="flex items-center gap-2">
                  <h3 className="font-black uppercase text-xs tracking-widest italic">Ficha de Remisión: {remisionSeleccionada.numero_remision}</h3>
                  
                  {saldoActual <= 0 ? (
                    <span className="bg-emerald-600 text-white px-2 py-0.5 rounded-md text-[9px] font-black uppercase">🟢 PAGADA</span>
                  ) : totalAbonado > 0 ? (
                    <span className="bg-amber-500 text-slate-900 px-2 py-0.5 rounded-md text-[9px] font-black uppercase">🟡 ABONADA</span>
                  ) : (
                    <span className="bg-rose-600 text-white px-2 py-0.5 rounded-md text-[9px] font-black uppercase">🔴 PENDIENTE DE PAGO</span>
                  )}
                </div>

                <span className="bg-green-700 px-2.5 py-0.5 rounded-md text-[9px] font-black uppercase">Detalle Carga</span>
              </div>

              <div className="p-5 space-y-5">
                <div className="grid grid-cols-1 md:grid-cols-2 gap-4 border-b pb-5">
                  <div>
                    <p className="text-[10px] font-black text-gray-400 uppercase italic">Cliente Destinatario</p>
                    <p className="font-black text-lg text-slate-900 uppercase leading-tight mt-1">{remisionSeleccionada.clientes?.nombre_completo}</p>
                    <p className="text-[10px] font-bold text-slate-400 mt-1 italic">Fecha Carga: {remisionSeleccionada.fecha_venta}</p>
                  </div>

                  <div className="bg-slate-50 p-4 rounded-2xl border border-slate-100">
                    <p className="text-[10px] font-black text-blue-600 uppercase mb-2 tracking-wider italic">Contenido Despachado</p>
                    <div className="space-y-1.5">
                      {remisionSeleccionada.detalle_ventas?.map((item, i) => (
                        <div key={i} className="flex justify-between items-center border-b border-dashed border-slate-200 pb-1">
                          <p className="text-[11px] font-bold text-slate-700 uppercase">{item.descripcion}</p>
                          <p className="text-[10px] font-black text-blue-700 bg-blue-100 px-2 py-0.5 rounded-md">
                            {item.amount || item.cantidad} {item.escala}
                          </p>
                        </div>
                      ))}
                      <div className="pt-2 flex justify-between items-center">
                        <p className="text-[9px] font-black text-slate-400 uppercase italic">Valor Total Venta:</p>
                        <p className="text-sm font-black text-green-700">{formatoPesos(remisionSeleccionada.total_venta)}</p>
                      </div>
                    </div>
                  </div>
                </div>

                {/* TABLA DISTRIBUIDA EN COLUMNAS */}
                <div className="space-y-3">
                  <p className="text-[10px] font-black text-slate-400 uppercase tracking-widest italic">Historial de Abonos Recibidos</p>
                  
                  {historialAbonos?.length > 0 ? (
                    <div className="overflow-x-auto rounded-2xl border border-slate-200 shadow-sm">
                      <table className="w-full text-left text-[11px] border-collapse">
                        <thead>
                          <tr className="bg-slate-100 text-slate-600 uppercase font-black border-b border-slate-200">
                            <th className="p-2.5 text-center">N°</th>
                            <th className="p-2.5">Fecha Abono</th>
                            <th className="p-2.5 text-center">Forma / Medio de Pago</th>
                            <th className="p-2.5">N° Ref / Comprobante</th>
                            <th className="p-2.5 text-right">Monto Abono</th>
                            <th className="p-2.5 text-center">Acciones</th>
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-slate-100 font-bold text-slate-700">
                          {historialAbonos.map((abono, idx) => {
                            const medioLimpio = obtenerMedioPagoLimpio(abono);

                            return (
                              <tr key={abono.id} className="hover:bg-sky-50/60 transition-colors">
                                <td className="p-2.5 text-center">
                                  <span className="bg-blue-100 text-blue-800 w-5 h-5 rounded-full inline-flex items-center justify-center font-black text-[9px]">
                                    {idx + 1}
                                  </span>
                                </td>

                                <td className="p-2.5 font-bold text-slate-800 whitespace-nowrap">
                                  {abono.fecha_pago}
                                </td>

                                <td className="p-2.5 text-center">
                                  <span className="bg-emerald-100 text-emerald-800 border border-emerald-300 px-2 py-0.5 rounded-md text-[9px] font-black uppercase tracking-tight inline-block">
                                    💳 {medioLimpio}
                                  </span>
                                </td>

                                <td className="p-2.5 uppercase text-slate-500 font-bold text-[10px]">
                                  {abono.referencia || abono.nota ? (
                                    <span className="bg-slate-100 px-2 py-0.5 rounded text-slate-700 font-black">
                                      {abono.referencia || abono.nota}
                                    </span>
                                  ) : (
                                    <span className="text-gray-300 italic">S/N</span>
                                  )}
                                </td>

                                <td className="p-2.5 text-right font-black text-blue-700 text-xs whitespace-nowrap">
                                  +{formatoPesos(abono.monto)}
                                </td>

                                <td className="p-2.5 text-center">
                                  <div className="flex gap-1 justify-center">
                                    <button 
                                      type="button"
                                      onClick={() => prepararEdicionPago(abono)}
                                      className="p-1 bg-amber-100 text-amber-700 rounded hover:bg-amber-600 hover:text-white transition-all text-xs"
                                      title="Editar"
                                    >
                                      ✏️
                                    </button>
                                    <button 
                                      type="button"
                                      onClick={() => eliminarPago(abono.id)}
                                      className="p-1 bg-red-100 text-red-700 rounded hover:bg-red-600 hover:text-white transition-all text-xs"
                                      title="Eliminar"
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
                    <div className="text-center py-5 bg-gray-50 rounded-xl border-2 border-dashed border-gray-200">
                      <p className="text-[10px] font-bold text-gray-400 uppercase italic">Sin pagos registrados a esta remisión</p>
                    </div>
                  )}
                </div>

                <div className="bg-slate-900 p-4 rounded-2xl flex justify-between items-center text-white shadow-md">
                  <div>
                    <p className="text-[9px] font-bold text-slate-400 uppercase">Total Recaudado</p>
                    <p className="font-black text-base text-blue-400">{formatoPesos(totalAbonado)}</p>
                  </div>
                  <div className="text-right">
                    <p className="text-[9px] font-bold text-slate-400 uppercase tracking-tighter">Saldo Neto Pendiente</p>
                    <p className={`font-black text-xl ${saldoActual <= 0 ? 'text-green-400' : 'text-red-400'}`}>
                      {formatoPesos(saldoActual)}
                    </p>
                  </div>
                </div>

                <div className="flex justify-end pt-1">
                  <button
                    type="button"
                    onClick={async () => {
                      await imprimirReciboCarteraPDF(remisionSeleccionada);
                    }}
                    className="px-4 py-2 bg-red-700 hover:bg-red-800 text-white font-black italic rounded-xl shadow-md transition-colors flex items-center gap-1.5 text-xs uppercase tracking-wider border border-red-600 cursor-pointer"
                  >
                    🖨️ Imprimir PDF Remisión
                  </button>
                </div>
              </div>
            </div>
          ) : (
            <div className="bg-blue-50 border-2 border-dashed border-blue-200 p-12 rounded-3xl text-center flex items-center justify-center h-full min-h-[300px]">
              <p className="text-blue-400 font-black uppercase text-xs italic">Seleccione un cliente y una remisión para cargar el desglose y su estado de cuenta</p>
            </div>
          )}
        </div>

      </div>
    </div>
  );
}