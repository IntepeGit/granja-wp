import { useEffect, useState } from 'react'
import { supabase } from './lib/supabase'
import { BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer } from 'recharts'
import jsPDF from 'jspdf'
import autoTable from 'jspdf-autotable' 

// 🛡️ Importación del componente de protección local (sin rutas)
import ProtectedRoute from './components/ProtectedRoute.jsx'

// Importaciones de componentes
import Login from './pages/Login.jsx'
import Despachos from './pages/Despachos.jsx' 
import Gastos from './pages/Gastos.jsx'
import Pagos from './pages/Pagos.jsx' 
import ConfigInv from './pages/ConfigInv.jsx'
import ConfigCli from './pages/ConfigCli.jsx'
import ConfigProv from './pages/ConfigProv.jsx'
import ReporteVentas from './pages/ReporteVentas.jsx'
import Dashboard from './pages/Dashboard.jsx'
import Inventario from './pages/Inventario.jsx'
import Cosecha from './pages/Cosecha.jsx'
import ConfigCosecha from './pages/ConfigCosecha.jsx'
import Nomina from './pages/Nomina.jsx'

function App() {
  const [session, setSession] = useState(null)
  const [userRole, setUserRole] = useState(null) // 🌟 NUEVO: Guarda 'admin' o 'operario'
  const [tab, setTab] = useState('dashboard')
  const [isMenuOpen, setIsMenuOpen] = useState(false)
  const [showConfigSubmenu, setShowConfigSubmenu] = useState(false)
  
  // Estados de datos consolidado
  const [datosDespachos, setDatosDespachos] = useState([]) 
  const [datosEgresos, setDatosEgresos] = useState([])
  const [datosPagos, setDatosPagos] = useState([]) 
  const [listaInvernaderos, setListaInvernaderos] = useState([])
  const [listaClientes, setListaClientes] = useState([])
  const [listaProveedores, setListaProveedores] = useState([])
  const [balancesGrafica, setBalancesGrafica] = useState([])
  const [notificacion, setNotificacion] = useState({ visible: false, mensaje: '', tipo: 'exito' });

  // Manejo de sesión original intacto
  useEffect(() => {
    const obtenerRolPerfil = async (userId) => {
      try {
        const { data, error } = await supabase
          .from('perfiles')
          .select('rol')
          .eq('id', userId)
          .single();
        
        if (!error && data) {
          setUserRole(data.rol); // Guarda 'admin' o 'operario' en la memoria
          if (data.rol === 'operario') {
            setTab('cosecha'); // Al operario lo manda directo a cosechas
          }
        }
      } catch (err) {
        console.error("Error obteniendo rol:", err);
      }
    };

    supabase.auth.getSession().then(({ data: { session } }) => {
      setSession(session);
      if (session?.user) obtenerRolPerfil(session.user.id);
    });

    const { data: { subscription } } = supabase.auth.onAuthStateChange((_event, session) => {
      setSession(session);
      if (session?.user) {
        obtenerRolPerfil(session.user.id);
      } else {
        setUserRole(null); // Limpia el rol si se sale del sistema
      }
    });

    return () => subscription.unsubscribe();
  }, []);

  const mostrarAlerta = (msj, tipo = 'exito') => {
    setNotificacion({ visible: true, mensaje: msj, tipo });
    setTimeout(() => setNotificacion(prev => ({ ...prev, visible: false })), 3000);
  };

  // ESTADOS DE FORMULARIOS
  const [despachoForm, setDespachoForm] = useState({ 
    numero_remision: '', 
    cliente_id: '', 
    invernadero_id: '', 
    fecha_venta: new Date().toISOString().split('T')[0], 
    filas: [{ producto: '', escala: '', cantidad: '', precio: '' }] 
  })
  
  const [gastoForm, setGastoForm] = useState({ id_editando: null, descripcion: '', categoria: 'Mano de obra', monto: '', invernadero_id: '', proveedor_id: '', numero_comprobante: '', nota: '', fecha: new Date().toISOString().split('T')[0], cantidad: '', precio_unitario: '', unidad_medida: 'Unidad' })
  const [invForm, setInvForm] = useState({id_editando: null, nombre: '', cultivo: '', variedad: '', largo: '', ancho: '', siembra: '', cosecha: '', estado: 'Activo', descripcion: ''})
  const [cliForm, setCliForm] = useState({ id_editando: null, nombre: '', nit: '', tel: '', dir: '', ciudad: '', nota: '', email: '' })
  const [provForm, setProvForm] = useState({ nombre: '', nit: '', tel: '', dir: '', ciudad: '', nota: '' })
  const [pagoForm, setPagoForm] = useState({ id_editando: null, cliente_id: '', venta_id: '', monto: '', metodo_pago: 'Efectivo', fecha_pago: new Date().toISOString().split('T')[0], nota: '' })
  
  const eliminarGasto = async (id) => {
    if (window.confirm("¿Está seguro de eliminar este gasto?")) {
      const { error } = await supabase.from('egresos').delete().eq('id', id);
      if (!error) { mostrarAlerta("Gasto eliminado"); cargarTodo(); }
    }
  };

  const prepararEdicionGasto = (g) => {
    setGastoForm({
      id_editando: g.id,
      descripcion: g.descripcion,
      categoria: g.categoria,
      monto: g.monto,
      invernadero_id: g.invernadero_id,
      proveedor_id: g.proveedor_id,
      numero_comprobante: g.numero_comprobante || '',
      nota: g.nota || '',
      fecha: g.fecha_gasto || g.fecha
    });
    window.scrollTo({ top: 0, behavior: 'smooth' });
  };

  const guardarGasto = async (e) => {
    e.preventDefault();
    const payload = {
      descripcion: gastoForm.descripcion,
      categoria: gastoForm.categoria,
      monto: parseFloat(gastoForm.monto),
      invernadero_id: gastoForm.invernadero_id,
      proveedor_id: gastoForm.proveedor_id,
      numero_comprobante: gastoForm.numero_comprobante,
      nota: gastoForm.nota,
      fecha_gasto: gastoForm.fecha
    };
    if (gastoForm.id_editando) {
      await supabase.from('egresos').update(payload).eq('id', gastoForm.id_editando);
    } else {
      await supabase.from('egresos').insert([payload]);
    }
    mostrarAlerta(gastoForm.id_editando ? "Gasto actualizado" : "Gasto guardado");
    setGastoForm({ id_editando: null, descripcion: '', categoria: 'Mano de obra', monto: '', invernadero_id: '', proveedor_id: '', numero_comprobante: '', nota: '', fecha: new Date().toISOString().split('T')[0] });
    cargarTodo();
  };

  const eliminarPago = async (id) => {
    if (window.confirm("¿Está seguro de eliminar este pago?")) {
      const { error } = await supabase.from('pagos').delete().eq('id', id);
      if (!error) { mostrarAlerta("Pago eliminado"); cargarTodo(); }
    }
  };

  const prepararEdicionPago = (pago) => {
    setPagoForm({
      id_editando: pago.id,
      cliente_id: pago.cliente_id,
      despacho_id: pago.despacho_id, 
      monto: pago.monto,
      fecha_pago: pago.fecha_pago,
      referencia: pago.referencia || '' 
    });
    window.scrollTo({ top: 0, behavior: 'smooth' });
  };

  const guardarPago = async (e) => {
    if (e && e.preventDefault) e.preventDefault();

    const payload = {
      cliente_id: pagoForm.cliente_id,
      despacho_id: pagoForm.despacho_id, 
      monto: parseFloat(pagoForm.monto),
      fecha_pago: pagoForm.fecha_pago,
      referencia: pagoForm.referencia 
    };

    try {
      if (pagoForm.id_editando) {
        await supabase.from('pagos').update(payload).eq('id', pagoForm.id_editando);
      } else {
        await supabase.from('pagos').insert([payload]);
      }

      mostrarAlerta(pagoForm.id_editando ? "Abono actualizado correctamente" : "Abono guardado");

      setPagoForm({ 
        ...pagoForm,           
        id_editando: null,     
        monto: '',             
        referencia: ''         
      });

      cargarTodo();
    } catch (error) {
      mostrarAlerta("Error: " + error.message, "error");
    }
  };

  const prepararEdicionDespacho = async (venta) => {
    try {
      const { data: detalles, error } = await supabase
        .from('detalle_ventas')
        .select('*')
        .eq('venta_id', venta.id);

      if (error) throw error;

      setDespachoForm({
        id_editando: venta.id,
        numero_remision: venta.numero_remision,
        cliente_id: venta.cliente_id,
        invernadero_id: venta.invernadero_id,
        fecha_venta: venta.fecha_venta,
        filas: detalles.map(d => ({
          producto: d.descripcion,
          escala: d.escala,
          cantidad: d.cantidad,
          precio: d.precio_unitario
        }))
      });
      
      window.scrollTo({ top: 0, behavior: 'smooth' });
      
    } catch (error) {
      mostrarAlerta("Error al cargar la remisión: " + error.message, "error");
    }
  };

  async function cargarTodo() {
    if (!session) return;
    
    const { data: v } = await supabase.from('ventas').select('*, clientes(*), invernaderos(*), detalle_ventas(*)').order('fecha_venta', { ascending: false });
    const { data: e } = await supabase.from('egresos').select('*, invernaderos(*), proveedores(*)');
    const { data: i } = await supabase.from('invernaderos').select('*');  
    const { data: c } = await supabase.from('clientes').select('*');
    const { data: p } = await supabase.from('proveedores').select('*');
    const { data: pagos } = await supabase.from('pagos').select('*, clientes(nombre_completo), ventas(numero_remision)');

    const invernaderosActivos = i?.filter(inv => inv.activo !== false) || [];

    setDatosDespachos(v || []);
    setDatosEgresos(e || []);
    setDatosPagos(pagos || []);
    setListaInvernaderos(i || []); 
    setListaClientes(c || []);
    setListaProveedores(p || []);

    const resumen = invernaderosActivos.map(inv => {
      const egr = e?.filter(x => x.invernadero_id === inv.id).reduce((s, x) => s + (x.monto || 0), 0) || 0;
      const ventasInv = v?.filter(x => x.invernadero_id === inv.id) || [];
      const ing = ventasInv.reduce((s, x) => s + (x.total_venta || 0), 0);
      
      const idsVentas = ventasInv.map(venta => venta.id);
      const pagosVentas = pagos?.filter(p => idsVentas.includes(p.despacho_id)) || [];
      const recaudado = pagosVentas.reduce((s, p) => s + (p.monto || 0), 0);
      const carteraPendiente = ing - recaudado;

      return { 
        name: inv.nombre, 
        Ingresos: ing, 
        Gastos: egr, 
        Utilidad: ing - egr,
        Cartera: carteraPendiente > 0 ? carteraPendiente : 0
      };
    });
    setBalancesGrafica(resumen || []);
  }  

  useEffect(() => {
    const validarRemisionEnVivo = async () => {
      if (despachoForm.numero_remision && !despachoForm.id_editando && tab === 'despachos') {
        const { data } = await supabase
          .from('ventas')
          .select('numero_remision')
          .eq('numero_remision', despachoForm.numero_remision.trim())
          .maybeSingle();

        if (data) {
          setDespachoForm(prev => ({ ...prev, errorDuplicado: true }));
        } else {
          setDespachoForm(prev => ({ ...prev, errorDuplicado: false }));
        }
      } else {
        setDespachoForm(prev => ({ ...prev, errorDuplicado: false }));
      }
    };

    const timeoutId = setTimeout(validarRemisionEnVivo, 500);
    return () => clearTimeout(timeoutId);
  }, [despachoForm.numero_remision, tab]);

  useEffect(() => { if (session) cargarTodo() }, [session]);

  useEffect(() => {
    if (tab === 'despachos' && !despachoForm.id_editando && datosDespachos.length > 0) {
      const numeros = datosDespachos
        .map(d => parseInt(d.numero_remision))
        .filter(n => !isNaN(n));
      
      const ultimoNumero = numeros.length > 0 ? Math.max(...numeros) : 0;
      
      if (!despachoForm.numero_remision) {
        setDespachoForm(prev => ({ 
          ...prev, 
          numero_remision: (ultimoNumero + 1).toString() 
        }));
      }
    }
  }, [tab, datosDespachos, despachoForm.id_editando]);

  const actualizarFilaDespacho = (index, campo, valor) => {
    const nuevasFilas = [...despachoForm.filas]
    nuevasFilas[index][campo] = valor
    setDespachoForm({ ...despachoForm, filas: nuevasFilas })
  }

  const guardarDespachoCompleto = async (e) => {
    if (e && e.preventDefault) e.preventDefault();
    
    try {
      const numRemision = despachoForm.numero_remision.trim();

      if (!despachoForm.id_editando) {
        const { data: existe } = await supabase
          .from('ventas')
          .select('numero_remision')
          .eq('numero_remision', numRemision);

        if (existe && existe.length > 0) {
          mostrarAlerta(`⚠️ ERROR: La remisión N° ${numRemision} ya existe en el historial.`, "error");
          return;
        }
      }

      const totalVentaCalculado = despachoForm.filas.reduce((acc, f) => 
        acc + (parseFloat(f.cantidad || 0) * parseFloat(f.precio || 0)), 0);

      const datosVenta = {
        numero_remision: numRemision,
        cliente_id: despachoForm.cliente_id,
        invernadero_id: despachoForm.invernadero_id,
        total_venta: totalVentaCalculado, 
        fecha_venta: despachoForm.fecha_venta
      };

      let ventaId = despachoForm.id_editando;

      if (ventaId) {
        const { error: vError } = await supabase.from('ventas').update(datosVenta).eq('id', ventaId);
        if (vError) throw vError;
        await supabase.from('detalle_ventas').delete().eq('venta_id', ventaId);
      } else {
        const { data: nuevaVenta, error: vError } = await supabase.from('ventas')
          .insert([datosVenta]).select().single();
        if (vError) throw vError;
        ventaId = nuevaVenta.id;
      }

      const detalles = despachoForm.filas.map(f => ({
        venta_id: ventaId,
        descripcion: f.producto,
        escala: f.escala,
        cantidad: parseFloat(f.cantidad || 0),
        precio_unitario: parseFloat(f.precio || 0)
      }));

      const { error: dError } = await supabase.from('detalle_ventas').insert(detalles);
      if (dError) throw dError;

      mostrarAlerta(despachoForm.id_editando ? "Remisión actualizada" : "Despacho guardado con éxito");
      
      setDespachoForm({ 
        id_editando: null, 
        numero_remision: '', 
        cliente_id: '', 
        invernadero_id: '', 
        fecha_venta: new Date().toISOString().split('T')[0], 
        filas: [{ producto: '', escala: '', cantidad: '', precio: '' }] 
      });
      
      await cargarTodo();

    } catch (error) {
      mostrarAlerta("Error técnico: " + error.message, "error");
    }
  };

  const eliminarDespacho = async (id) => {
    if (window.confirm("¿Está seguro de eliminar esta remisión?")) {
      await supabase.from('detalle_ventas').delete().eq('venta_id', id);
      const { error } = await supabase.from('ventas').delete().eq('id', id);
      if (!error) {
        mostrarAlerta("Remisión Doc. Eliminada");
        cargarTodo();
      }
    }
  };

  const imprimirPDF = (remision) => {
    try {
      const detalles = remision.detalle_ventas || [];
      if (detalles.length === 0) {
        mostrarAlerta("Esta remisión no tiene productos", "error");
        return;
      }
      const doc = new jsPDF({ orientation: 'portrait', unit: 'mm', format: [105, 148] });
      // Bordes en el azul del invernadero (17, 112, 151)
      doc.setDrawColor(17, 112, 151); doc.setLineWidth(0.8); doc.rect(4, 4, 97, 140); 
      doc.addImage('/Logopapel.png', 'PNG', 42.5, 7, 20, 20); 
      doc.setFont("helvetica", "bold"); doc.setFontSize(9); doc.text("N° Remisión:", 10, 12);
      doc.setFont("helvetica", "normal"); doc.text(`${remision.numero_remision || 'N/A'}`, 32, 12);
      doc.setFont("helvetica", "bold"); doc.setFontSize(11); doc.setDrawColor(17, 112, 151);
      doc.text("REMISIÓN DE VENTA", 52.5, 30, { align: "center" });

      const altoFila = 5; const yBase = 32.5;
      // Fondo muy suave azul-celeste (215, 235, 245)
      doc.setFillColor(215, 235, 245); doc.rect(7, yBase, 91, altoFila, 'F'); doc.rect(7, yBase + (altoFila * 2), 91, altoFila, 'F');
      doc.setLineWidth(0.2); doc.rect(7, yBase, 91, altoFila * 4);
      doc.line(7, yBase + altoFila, 98, yBase + altoFila); doc.line(7, yBase + (altoFila * 2), 98, yBase + (altoFila * 2)); doc.line(7, yBase + (altoFila * 3), 98, yBase + (altoFila * 3));
      doc.line(50, yBase, 50, yBase + (altoFila * 4));

      const yOffset = 3.5;
      doc.setFont("helvetica", "bold"); doc.text("Fecha:", 10, yBase + yOffset);
      doc.setFont("helvetica", "normal"); doc.text(`${remision.fecha_venta || ''}`, 32, yBase + yOffset);
      doc.setFont("helvetica", "bold"); doc.text("Ciudad:", 52, yBase + yOffset);
      doc.setFont("helvetica", "normal"); doc.text(`${remision.clientes?.ciudad || 'N/A'}`, 68, yBase + yOffset);
      doc.setFont("helvetica", "bold"); doc.text("Invernadero:", 10, yBase + altoFila + yOffset);
      doc.setFont("helvetica", "normal"); doc.text(`${remision.invernaderos?.nombre || 'N/A'}`, 32, yBase + altoFila + yOffset);
      doc.setFont("helvetica", "bold"); doc.text("Celular:", 52, yBase + altoFila + yOffset);
      doc.setFont("helvetica", "normal"); doc.text(`${remision.clientes?.telefono || 'N/A'}`, 68, yBase + altoFila + yOffset);
      doc.setFont("helvetica", "bold"); doc.text("Cliente:", 10, yBase + (altoFila * 2) + yOffset);
      doc.setFont("helvetica", "normal"); doc.text(`${remision.clientes?.nombre_completo || 'N/A'}`, 32, yBase + (altoFila * 2) + yOffset);
      doc.setFont("helvetica", "bold"); doc.text("Correo:", 52, yBase + (altoFila * 2) + yOffset);
      doc.setFontSize(5.5); doc.text(`${remision.clientes?.email || 'N/A'}`, 68, yBase + (altoFila * 2) + yOffset);
      doc.setFontSize(8); doc.setFont("helvetica", "bold"); doc.text("NIT/CC:", 10, yBase + (altoFila * 3) + yOffset);
      doc.setFont("helvetica", "normal"); doc.text(`${remision.clientes?.nit || 'N/A'}`, 32, yBase + (altoFila * 3) + yOffset);
      
      autoTable(doc, {
        startY: 56, head: [["Cant", "Producto", "Precio", "Total"]],
        body: detalles.map(item => [
          item.cantidad || 0, item.descripcion || 'Sin desc.',
          new Intl.NumberFormat('es-CO').format(item.precio_unitario || 0),
          new Intl.NumberFormat('es-CO').format((item.cantidad || 0) * (item.precio_unitario || 0))
        ]),
        theme: 'grid', styles: { fontSize: 7, cellPadding: 1.2, fontStyle: 'bold' },
        // Cabecera en azul principal
        headStyles: { fillColor: [17, 112, 151], textColor: [255, 255, 255] }
      });

      const finalY = doc.lastAutoTable.finalY + 8;
      doc.setFontSize(10); doc.setFont("helvetica", "bold");
      doc.text(`TOTAL A PAGAR: ${new Intl.NumberFormat('es-CO', { style: 'currency', currency: 'COP' }).format(remision.total_venta || 0)}`, 95, finalY, { align: "right" });
      doc.save(`Remision_${remision.numero_remision}.pdf`);
    } catch (err) {
      console.error(err);
    }
  };

  const imprimirGastoPDF = (gasto) => {
    try {
      // Crear documento A6 idéntico a tu configuración
      const doc = new jsPDF({ orientation: 'portrait', unit: 'mm', format: [105, 148] });
      
      // 🫐 Borde exterior decorativo en Azul Océano
      doc.setDrawColor(17, 112, 151); 
      doc.setLineWidth(0.8); 
      doc.rect(4, 4, 97, 140);

      // 🫐 Imagen del logo centrada
      try {
        doc.addImage('/Logopapel.png', 'PNG', 42.5, 7, 20, 20);
      } catch (e) {
        console.log("No se pudo cargar el logo en el PDF");
      }

      // 🫐 Consecutivo del Comprobante (Arriba a la izquierda)
      doc.setFont("helvetica", "bold"); 
      doc.setFontSize(8); 
      doc.setTextColor(60, 60, 60);
      doc.text(`Comp. N°: ${gasto.numero_comprobante || gasto.id || 'S/N'}`, 7, 12);

      // 🫐 Título Principal
      doc.setFont("helvetica", "bold"); 
      doc.setFontSize(11); 
      doc.setTextColor(17, 112, 151);
      doc.text("COMPROBANTE DE GASTO", 52.5, 31, { align: "center" });

      // 🗃️ SOLUCIÓN MAESTRA: Datos de cabecera organizados en una AutoTable invisible y compacta
      const nombreProv = gasto.nombre_proveedor || gasto.proveedores?.nombre_completo || gasto.proveedores?.nombre || 'Particular';
      const nit = gasto.nit_cc || gasto.proveedores?.nit_cc || 'N/A';
      const tel = gasto.telefono_proveedor || gasto.proveedores?.telefono || 'N/R';
      const ciudad = gasto.ciudad_proveedor || gasto.proveedores?.ciudad || 'N/R';
      const invernadero = gasto.nombre_invernadero || gasto.invernaderos?.nombre || 'General';

      autoTable(doc, {
        startY: 35,
        margin: { left: 6, right: 6 },
        theme: 'plain', // Sin bordes ni fondos toscos
        styles: { fontSize: 7.5, cellPadding: 1, font: "helvetica" },
        columnStyles: {
          0: { cellWidth: 18, fontStyle: 'bold', textColor: [17, 112, 151] }, // Etiquetas en azul
          1: { cellWidth: 28, textColor: [40, 40, 40] },
          2: { cellWidth: 14, fontStyle: 'bold', textColor: [17, 112, 151] },
          3: { cellWidth: 33, textColor: [40, 40, 40] }
        },
        body: [
          ["Fecha:", `${gasto.fecha_gasto || ''}`, "Lote/Inv:", `${invernadero.toUpperCase()}`],
          ["Proveedor:", `${nombreProv.toUpperCase()}`, "NIT/CC:", `${nit}`],
          ["Teléfono:", `${tel}`, "Ciudad:", `${ciudad.toUpperCase()}`],
          ["Nota/Obs:", { content: `${gasto.nota || 'Sin observaciones.'}`, colSpan: 3, styles: { fontStyle: 'italic' } }]
        ]
      });

      // 🫐 Tabla Inferior de los Ítems del Gasto (Se posiciona automáticamente abajo de la cabecera)
      const yTablaDetalle = doc.lastAutoTable.finalY + 3;
      autoTable(doc, {
        startY: yTablaDetalle, 
        margin: { left: 6, right: 6 }, 
        head: [["Cant.", "Unidad", "Detalle del Pago", "Precio Unit.", "Monto Total"]],
        body: [[
          gasto.cantidad || 0, 
          gasto.unidad_medida || "Unidad", 
          gasto.descripcion || "Gasto", 
          `$${(gasto.precio_unitario || 0).toLocaleString('es-CO')}`, 
          `$${(gasto.monto || 0).toLocaleString('es-CO')}`
        ]],
        theme: 'grid', 
        styles: { fontSize: 7, cellPadding: 2, valign: 'middle', overflow: 'linebreak', lineWidth: 0.1, lineColor: [17, 112, 151] },
        headStyles: { fillColor: [17, 112, 151], textColor: [255, 255, 255], halign: 'center', fontStyle: 'bold' },
        columnStyles: { 
          0: { cellWidth: 9, halign: 'center' }, 
          1: { cellWidth: 14, halign: 'center' }, 
          2: { cellWidth: 35 }, 
          3: { cellWidth: 16, halign: 'right' }, 
          4: { cellWidth: 19, halign: 'right' } 
        }
      });

      // 🫐 Bloque de Totales Consolidado
      const finalY = doc.lastAutoTable.finalY + 5;
      doc.setFontSize(9.5); 
      doc.setFont("helvetica", "bold");
      doc.setTextColor(0);
      doc.text(`VALOR PAGADO: ${new Intl.NumberFormat('es-CO', { style: 'currency', currency: 'COP', maximumFractionDigits: 0 }).format(gasto.monto || 0)}`, 99, finalY, { align: "right" });

      // 🫐 Sección de Firmas de Control inferior
      const yFirmas = Math.max(finalY + 15, 128); // Asegura que las firmas se queden abajo del papel A6
      doc.setDrawColor(180, 180, 180);
      doc.setLineWidth(0.3);
      doc.line(8, yFirmas, 46, yFirmas); 
      doc.line(59, yFirmas, 97, yFirmas);
      doc.setFontSize(6.5); 
      doc.text("ENTREGUÉ CONFORME", 27, yFirmas + 3.5, { align: "center" }); 
      doc.text("RECIBÍ CONFORME", 78, yFirmas + 3.5, { align: "center" });

      // Guardar archivo
      doc.save(`Gasto_${gasto.numero_comprobante || gasto.id}.pdf`);
    } catch (err) {
      console.error("Error al generar PDF de Gasto:", err);
    }
  };

  const prepararEdicion = (despacho) => {
    setDespachoForm({
      id_editando: despacho.id,
      numero_remision: despacho.numero_remision,
      cliente_id: despacho.cliente_id,
      invernadero_id: despacho.invernadero_id,
      fecha_venta: despacho.fecha_venta,
      filas: despacho.detalle_ventas.map(d => ({
        producto: d.descripcion, escala: d.escala, cantidad: d.cantidad, precio: d.precio_unitario
      }))
    });
    window.scrollTo({ top: 0, behavior: 'smooth' });
  };

  // 🫐 NavItem unificado con tu color exacto (#3B4E38) para inactivos y variantes contrastantes
  const NavItem = ({ id, label, icon }) => (
    <button onClick={() => { setTab(id); setIsMenuOpen(false); setShowConfigSubmenu(false); }} 
      className={`flex items-center gap-3 w-full p-4 rounded-xl transition ${tab === id ? 'bg-[#0a4c68] text-white shadow-inner font-black' : 'text-[#b6e2f8] hover:bg-[#0a4c68] hover:text-white'}`}>
      <span className="text-xl">{icon}</span> <span className="font-bold text-sm capitalize">{label}</span>
    </button>
  )

  // 🛡️ CONTROL DE ACCESO INTERNO (Si no hay sesión, renderiza el Login de inmediato)
  if (!session) return <Login setSession={setSession} />;

  return (
    <ProtectedRoute session={session}>
      <div className="flex min-h-screen bg-slate-50 font-sans overflow-hidden">
        {/* 🫐 Sidebar unificado exactamente a tu color de imagen (#3B4E38) */}
        <aside className={`fixed inset-y-0 left-0 z-50 w-64 bg-[#3B4E38] shadow-2xl transform ${isMenuOpen ? 'translate-x-0' : '-translate-x-full'} lg:translate-x-0 lg:static transition-transform duration-300 flex flex-col`}>
  
            {/* 🎯 Encabezado con bordes sutiles en azul más oscuro para dar profundidad */}
            <div className="p-6 text-center border-b border-[#0d5978] flex flex-col items-center justify-center gap-2">
              <h2 className="text-white font-black text-2xl tracking-tighter">🫐 INVERNADERO FM</h2>
              
              {userRole ? (
                <span className={`inline-flex items-center px-3 py-1 rounded-full text-[10px] font-black uppercase tracking-widest border shadow-inner ${
                  userRole === 'admin' 
                    ? 'bg-amber-500/20 text-amber-300 border-amber-500/30' 
                    : 'bg-[#0a4c68] text-[#d6eefc] border-[#083a50]'
                }`}>
                  {userRole === 'admin' ? '👑 Administrador' : '👨‍🌾 Operario'}
                </span>
              ) : (
                <span className="text-[10px] text-[#b6e2f8] animate-pulse font-bold uppercase tracking-widest">Cargando perfil...</span>
              )}
            </div>
          <nav className="p-4 space-y-2 flex-1 overflow-y-auto">
            
            {/* 👑 SOLO ADMIN: Dashboard de Balances */}
            {userRole === 'admin' && <NavItem id="dashboard" label="Dashboard" icon="📊" />}
            
            {/* 🚛 AMBOS: Despachos de camiones y carga */}
            <NavItem id="despachos" label="Despachos" icon="🚛" />
            
            {/* 👑 SOLO ADMIN: Cuentas por cobrar y abonos */}
            {userRole === 'admin' && <NavItem id="pagos" label="Pagos / Caja" icon="💳" />}
            
            {/* 👑 SOLO ADMIN: Costos operativos y facturas de egresos */}
            {userRole === 'admin' && <NavItem id="gastos" label="Gastos" icon="📉" />}
            
            {/* 👑 SOLO ADMIN: Auditorías de rendimiento comercial */}
            {userRole === 'admin' && <NavItem id="reporte" label="Reporte de Ventas" icon="📈" />}
            
            {/* 🌾 AMBOS: Lógica de recolección diaria de la tomatera */}
            <NavItem id="inventario" label="Inventario Bodega" icon="📦" />
            <NavItem id="cosecha" label="Cosecha Diaria" icon="🚜" />
            
            {/* 👑 SOLO ADMIN: Control de jornales y pagos a trabajadores */}
            {userRole === 'admin' && <NavItem id="nomina" label="Nómina / Mano Obra" icon="👥" />}
            
            {/* 👑 SOLO ADMIN: Parámetros del sistema y catálogos */}
            {userRole === 'admin' && (
              <div className="space-y-1">
                <button onClick={() => setShowConfigSubmenu(!showConfigSubmenu)} 
                  className={`flex items-center justify-between w-full p-4 rounded-xl transition ${tab.startsWith('config-') ? 'bg-[#0a4c68] text-white' : 'text-[#b6e2f8] hover:bg-[#0a4c68] hover:text-white'}`}>
                  <div className="flex items-center gap-3">
                    <span className="text-xl">⚙️</span>
                    <span className="font-bold text-sm">Configuración</span>
                  </div>
                  <span className="text-xs">{showConfigSubmenu ? '▲' : '▼'}</span>
                </button>
                {showConfigSubmenu && (
                  <div className="pl-6 space-y-1 animate-in slide-in-from-top-2 duration-200">
                    <button onClick={() => { setTab('config-inv'); setIsMenuOpen(false); }} className={`flex items-center gap-3 w-full p-3 rounded-lg text-xs font-bold transition ${tab === 'config-inv' ? 'text-white bg-[#0a4c68]' : 'text-[#b6e2f8] hover:bg-[#0a4c68]'}`}>🏠 Invernaderos</button>
                    <button onClick={() => { setTab('config-cli'); setIsMenuOpen(false); }} className={`flex items-center gap-3 w-full p-3 rounded-lg text-xs font-bold transition ${tab === 'config-cli' ? 'text-white bg-[#0a4c68]' : 'text-[#b6e2f8] hover:bg-[#0a4c68]'}`}>👥 Clientes</button>
                    <button onClick={() => { setTab('config-prov'); setIsMenuOpen(false); }} className={`flex items-center gap-3 w-full p-3 rounded-lg text-xs font-bold transition ${tab === 'config-prov' ? 'text-white bg-[#0a4c68]' : 'text-[#b6e2f8] hover:bg-[#0a4c68]'}`}>🚚 Proveedores</button>
                    <button onClick={() => { setTab('config-cosecha'); setIsMenuOpen(false); }} className={`flex items-center gap-3 w-full p-3 rounded-lg text-xs font-bold transition ${tab === 'config-cosecha' ? 'text-white bg-[#0a4c68]' : 'text-[#b6e2f8] hover:bg-[#0a4c68]'}`}>🌿 Parámetros Cosecha</button>
                  </div>
                )}
              </div>
            )}
            
            <button onClick={() => supabase.auth.signOut()} className="flex items-center gap-3 w-full p-4 rounded-xl text-red-200 hover:bg-red-950/20 mt-10">
              <span>🚪</span> <span className="text-sm font-bold">Cerrar Sesión</span>
            </button>
          </nav>
        </aside>

        <div className="flex-1 flex flex-col h-screen overflow-hidden">
          {/* Header con títulos en azul oscuro de alto contraste `text-[#0a4c68]` */}
          <header className="bg-white p-4 shadow-sm flex justify-between items-center lg:px-10">
            <button className="lg:hidden text-2xl p-2 text-[#0a4c68]" onClick={() => setIsMenuOpen(true)}>☰</button>
            <h1 className="text-xl font-black text-[#0a4c68] tracking-tight uppercase">{tab.replace('config-', 'Configuración: ').replace('reporte', 'Reporte de Ventas').replace('inventario', 'Inventario de Bodega').replace('cosecha', 'Cosecha Diaria').replace('nomina', 'Control de Nómina')}</h1>
            <div className="w-10"></div>
          </header>

          <main className="flex-1 overflow-y-auto p-4 md:p-10 space-y-10">
          
            {tab === 'dashboard' && (
              <Dashboard 
                listaInvernaderos={listaInvernaderos}
                datosDespachos={datosDespachos}
                datosEgresos={datosEgresos}
                datosPagos={datosPagos}
                balancesGrafica={balancesGrafica}
              />
            )}

            {tab === 'despachos' && (
              <Despachos 
                despachoForm={despachoForm}
                setDespachoForm={setDespachoForm} 
                listaClientes={listaClientes} 
                listaInvernaderos={listaInvernaderos.filter(inv => inv.activo !== false)}
                actualizarFilaDespacho={actualizarFilaDespacho}
                guardarDespachoCompleto={guardarDespachoCompleto} 
                datosDespachos={datosDespachos} 
                datosPagos={datosPagos} 
                mostrarAlerta={mostrarAlerta} 
                guardarDespacho={guardarDespachoCompleto}
                eliminarDespacho={eliminarDespacho} 
                prepararEdicion={prepararEdicion}
                prepararEdicionDespacho={prepararEdicionDespacho}
                imprimirPDF={imprimirPDF}
                cargarTodo={cargarTodo}
                supabase={supabase}
                userRole={userRole}
              />
            )}

            {tab === 'pagos' && (
              <Pagos 
                listaClientes={listaClientes} 
                datosDespachos={datosDespachos} 
                guardarPago={guardarPago} 
                datosPagos={datosPagos} 
                mostrarAlerta={mostrarAlerta} 
                pagoForm={pagoForm}
                setPagoForm={setPagoForm}
                cargarTodo={cargarTodo} 
                prepararEdicionPago={prepararEdicionPago}
                eliminarPago={eliminarPago}
              />
            )}
                      
            {tab === 'gastos' && (
              <Gastos 
                gastoForm={gastoForm}
                setGastoForm={setGastoForm} 
                listaInvernaderos={listaInvernaderos.filter(inv => inv.activo !== false)}
                listaProveedores={listaProveedores} 
                mostrarAlerta={mostrarAlerta}
                cargarTodo={cargarTodo} 
                supabase={supabase}
                datosEgresos={datosEgresos}
                guardarGasto={guardarGasto}
                prepararEdicionGasto={prepararEdicionGasto}
                eliminarGasto={eliminarGasto}
                imprimirGastoPDF={imprimirGastoPDF}
              />
            )}  

            {tab === 'reporte' && (
              <ReporteVentas 
                listaInvernaderos={listaInvernaderos}
                datosDespachos={datosDespachos}
                datosEgresos={datosEgresos}
                datosPagos={datosPagos}
                balancesGrafica={balancesGrafica}
              />
            )}

            {tab === 'inventario' && (
              <Inventario 
                mostrarAlerta={mostrarAlerta}
                datosInvernaderos={listaInvernaderos}
                userRole={userRole}
              />
            )}

            {tab === 'cosecha' && (
              <Cosecha 
                mostrarAlerta={mostrarAlerta}
                listaInvernaderos={listaInvernaderos}
                userRole={userRole}
              />
            )}

            {tab === 'nomina' && (
              <Nomina 
                mostrarAlerta={mostrarAlerta}
                listaInvernaderos={listaInvernaderos}
              />
            )}

            {tab === 'config-inv' && <ConfigInv invForm={invForm} setInvForm={setInvForm} mostrarAlerta={mostrarAlerta} cargarTodo={cargarTodo} supabase={supabase} lista={listaInvernaderos} />}
            {tab === 'config-cli' && <ConfigCli cliForm={cliForm} setCliForm={setCliForm} mostrarAlerta={mostrarAlerta} cargarTodo={cargarTodo} supabase={supabase} lista={listaClientes} />}
            {tab === 'config-prov' && <ConfigProv provForm={provForm} setProvForm={setProvForm} mostrarAlerta={mostrarAlerta} cargarTodo={cargarTodo} supabase={supabase} lista={listaProveedores} />}
          
            {tab === 'config-cosecha' && (
              <ConfigCosecha 
                mostrarAlerta={mostrarAlerta} 
              />
            )}
          
          </main>
        </div>

        {/* Notificación en azul con bordes coordinados */}
        {notificacion.visible && (
          <div className="fixed bottom-10 right-10 z-[100] bg-white border border-[#3B4E38]/20 p-6 rounded-2xl shadow-2xl flex items-center gap-3 animate-in slide-in-from-bottom-5">
            <span className="text-2xl">{notificacion.tipo === 'exito' ? '✅' : '❌'}</span>
            <p className="text-sm font-bold text-[#3B4E38]">{notificacion.mensaje}</p>
          </div>
        )}
      </div>
    </ProtectedRoute>
  )
}

export default App;