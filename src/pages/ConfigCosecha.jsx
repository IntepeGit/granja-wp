import React, { useState, useEffect } from 'react';
import { supabase } from '../lib/supabase';

export default function ConfigCosecha({ mostrarAlerta }) {
  const [productos, setProductos] = useState([]);
  const [calidades, setCalidades] = useState([]);
  const [unidades, setUnidades] = useState([]);
  const [cargando, setCargando] = useState(false);

  const [nuevoProducto, setNuevoProducto] = useState('');
  const [nuevaCalidad, setNuevaCalidad] = useState('');
  const [nuevaUnidad, setNuevaUnidad] = useState('');

  useEffect(() => {
    cargarTodosLosParametros();
  }, []);

  const cargarTodosLosParametros = async () => {
    setCargando(true);
    try {
      const [resProd, resCal, resUni] = await Promise.all([
        supabase.from('config_productos').select('*').order('nombre_producto', { ascending: true }),
        supabase.from('config_calidades').select('*').order('nombre_calidad', { ascending: true }),
        supabase.from('config_unidades').select('*').order('nombre_unidad', { ascending: true })
      ]);

      if (resProd.error) throw resProd.error;
      if (resCal.error) throw resCal.error;
      if (resUni.error) throw resUni.error;

      setProductos(resProd.data || []);
      setCalidades(resCal.data || []);
      setUnidades(resUni.data || []);
    } catch (err) {
      console.error("Error cargando parámetros:", err);
      mostrarAlerta("No se pudieron cargar las configuraciones", "error");
    } finally {
      setCargando(false);
    }
  };

  const agregarProducto = async (e) => {
    e.preventDefault();
    if (!nuevoProducto.trim()) return;
    try {
      const { error } = await supabase
        .from('config_productos')
        .insert([{ nombre_producto: nuevoProducto.trim().toUpperCase() }]);
      
      if (error) throw error;
      mostrarAlerta("Producto añadido exitosamente", "exito");
      setNuevoProducto('');
      cargarTodosLosParametros();
    } catch (err) {
      mostrarAlerta("El producto ya existe o hubo un error", "error");
    }
  };

  const agregarCalidad = async (e) => {
    e.preventDefault();
    if (!nuevaCalidad.trim()) return;
    try {
      const { error } = await supabase
        .from('config_calidades')
        .insert([{ nombre_calidad: nuevaCalidad.trim().toUpperCase() }]);
      
      if (error) throw error;
      mostrarAlerta("Calidad añadida exitosamente", "exito");
      setNuevaCalidad('');
      cargarTodosLosParametros();
    } catch (err) {
      mostrarAlerta("La calidad ya existe o hubo un error", "error");
    }
  };

  const agregarUnidad = async (e) => {
    e.preventDefault();
    if (!nuevaUnidad.trim()) return;
    try {
      const { error } = await supabase
        .from('config_unidades')
        .insert([{ nombre_unidad: nuevaUnidad.trim().toUpperCase() }]);
      
      if (error) throw error;
      mostrarAlerta("Unidad de medida añadida exitosamente", "exito");
      setNuevaUnidad('');
      cargarTodosLosParametros();
    } catch (err) {
      mostrarAlerta("La unidad ya existe o hubo un error", "error");
    }
  };

  const eliminarElemento = async (tabla, id, campoNombre, valorNombre) => {
    if (!window.confirm(`¿Seguro que deseas eliminar "${valorNombre}" de las opciones?`)) return;
    try {
      const { error } = await supabase.from(tabla).delete().eq('id', id);
      if (error) throw error;
      mostrarAlerta(`"${valorNombre}" eliminado con éxito`, "exito");
      cargarTodosLosParametros();
    } catch (err) {
      mostrarAlerta("No se pudo eliminar. Puede estar en uso en el historial de cosechas.", "error");
    }
  };

  return (
    <div className="space-y-4 pb-10 text-slate-800 dark:text-slate-200 font-sans transition-colors duration-300">
      
      {/* 🚀 CABECERA PRINCIPAL MODERNA */}
      <div className="bg-white dark:bg-slate-800 p-5 rounded-3xl shadow-xl border border-slate-200 dark:border-slate-700 flex flex-col md:flex-row justify-between items-center gap-4 transition-colors duration-300">
        <div className="space-y-0.5">
          <div className="flex items-center gap-3">
            <span className="text-xl p-2 bg-emerald-700/10 dark:bg-emerald-500/20 rounded-xl text-emerald-700 dark:text-emerald-400">⚙️</span>
            <h2 className="text-lg font-black tracking-tight text-slate-900 dark:text-white uppercase">Parámetros de Cosecha</h2>
          </div>
          <p className="text-xs text-slate-500 dark:text-slate-400 font-medium">Administración de catálogos, clasificaciones y unidades de medida.</p>
        </div>
      </div>

      {/* 💡 BANNER DE INFORMACIÓN */}
      <div className="bg-amber-50 dark:bg-amber-950/40 p-3 rounded-2xl shadow-sm border border-amber-200 dark:border-amber-800 flex gap-3 items-center transition-colors duration-300">
        <span className="text-2xl">💡</span>
        <div>
          <p className="text-[11px] font-black text-amber-800 dark:text-amber-400 uppercase tracking-wider">Información de Configuración</p>
          <p className="text-[10px] text-amber-700/80 dark:text-amber-300/80 font-bold">
            Los elementos que agregues aquí aparecerán de forma inmediata como opciones desplegables dentro del formulario principal de <span className="font-black underline">Cosecha Diaria</span>.
          </p>
        </div>
      </div>

      {/* 📊 CONTENEDOR DE LAS 3 COLUMNAS EXPANDIDO Y CON SCROLL */}
      <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
        
        {/* COLUMNA 1: PRODUCTOS */}
        <div className="bg-white dark:bg-slate-800 rounded-3xl shadow-xl border border-slate-200 dark:border-slate-700 overflow-hidden flex flex-col h-[calc(100vh-270px)] min-h-[400px] transition-colors duration-300">
          <div className="p-3 bg-slate-800 dark:bg-slate-900 text-white font-black text-[11px] uppercase tracking-wider flex items-center gap-2">
            <span>🍅</span> Catálogo de Productos
          </div>
          <div className="p-3.5 flex-1 flex flex-col space-y-3 min-h-0">
            <form onSubmit={agregarProducto} className="flex gap-2 shrink-0">
              <input 
                type="text" 
                className="flex-1 bg-slate-50 dark:bg-slate-900 border border-slate-200 dark:border-slate-700 text-slate-800 dark:text-white px-3 py-2 rounded-xl font-bold text-xs uppercase outline-none focus:border-emerald-700 transition-colors placeholder-slate-400" 
                value={nuevoProducto} 
                onChange={e => setNuevoProducto(e.target.value)} 
                placeholder="Nuevo Producto..." 
                required 
              />
              <button type="submit" className="px-3.5 bg-emerald-600 hover:bg-emerald-700 text-white font-black rounded-xl text-base shadow-md transition-colors cursor-pointer">
                ＋
              </button>
            </form>
            
            <div className="flex-1 overflow-y-auto pr-1 space-y-1.5 custom-scrollbar min-h-0">
              {productos.length === 0 ? (
                <p className="text-center text-[10px] font-bold text-slate-400 dark:text-slate-500 italic mt-4">No hay productos registrados.</p>
              ) : (
                productos.map(p => (
                  <div key={p.id} className="px-3 py-2 flex justify-between items-center text-[11px] font-black uppercase text-slate-700 dark:text-slate-200 bg-slate-50 dark:bg-slate-900/60 border border-slate-100 dark:border-slate-700 rounded-xl hover:border-emerald-200 transition-colors">
                    <span>{p.nombre_producto}</span>
                    <button 
                      onClick={() => eliminarElemento('config_productos', p.id, 'nombre_producto', p.nombre_producto)} 
                      className="p-1 bg-red-100 dark:bg-red-950/60 text-red-700 dark:text-red-400 hover:bg-red-600 hover:text-white rounded-lg transition-colors cursor-pointer text-xs"
                      title="Eliminar Producto"
                    >
                      🗑️
                    </button>
                  </div>
                ))
              )}
            </div>
          </div>
        </div>

        {/* COLUMNA 2: CALIDADES */}
        <div className="bg-white dark:bg-slate-800 rounded-3xl shadow-xl border border-slate-200 dark:border-slate-700 overflow-hidden flex flex-col h-[calc(100vh-270px)] min-h-[400px] transition-colors duration-300">
          <div className="p-3 bg-slate-800 dark:bg-slate-900 text-white font-black text-[11px] uppercase tracking-wider flex items-center gap-2">
            <span>⭐</span> Clasificación / Calidad
          </div>
          <div className="p-3.5 flex-1 flex flex-col space-y-3 min-h-0">
            <form onSubmit={agregarCalidad} className="flex gap-2 shrink-0">
              <input 
                type="text" 
                className="flex-1 bg-slate-50 dark:bg-slate-900 border border-slate-200 dark:border-slate-700 text-slate-800 dark:text-white px-3 py-2 rounded-xl font-bold text-xs uppercase outline-none focus:border-emerald-700 transition-colors placeholder-slate-400" 
                value={nuevaCalidad} 
                onChange={e => setNuevaCalidad(e.target.value)} 
                placeholder="Nueva Calidad..." 
                required 
              />
              <button type="submit" className="px-3.5 bg-emerald-600 hover:bg-emerald-700 text-white font-black rounded-xl text-base shadow-md transition-colors cursor-pointer">
                ＋
              </button>
            </form>

            <div className="flex-1 overflow-y-auto pr-1 space-y-1.5 custom-scrollbar min-h-0">
              {calidades.length === 0 ? (
                <p className="text-center text-[10px] font-bold text-slate-400 dark:text-slate-500 italic mt-4">No hay calidades registradas.</p>
              ) : (
                calidades.map(c => (
                  <div key={c.id} className="px-3 py-2 flex justify-between items-center text-[11px] font-black uppercase text-slate-700 dark:text-slate-200 bg-slate-50 dark:bg-slate-900/60 border border-slate-100 dark:border-slate-700 rounded-xl hover:border-emerald-200 transition-colors">
                    <span>{c.nombre_calidad}</span>
                    <button 
                      onClick={() => eliminarElemento('config_calidades', c.id, 'nombre_calidad', c.nombre_calidad)} 
                      className="p-1 bg-red-100 dark:bg-red-950/60 text-red-700 dark:text-red-400 hover:bg-red-600 hover:text-white rounded-lg transition-colors cursor-pointer text-xs"
                      title="Eliminar Calidad"
                    >
                      🗑️
                    </button>
                  </div>
                ))
              )}
            </div>
          </div>
        </div>

        {/* COLUMNA 3: UNIDADES DE MEDIDA */}
        <div className="bg-white dark:bg-slate-800 rounded-3xl shadow-xl border border-slate-200 dark:border-slate-700 overflow-hidden flex flex-col h-[calc(100vh-270px)] min-h-[400px] transition-colors duration-300">
          <div className="p-3 bg-slate-800 dark:bg-slate-900 text-white font-black text-[11px] uppercase tracking-wider flex items-center gap-2">
            <span>⚖️</span> Unidades de Medida
          </div>
          <div className="p-3.5 flex-1 flex flex-col space-y-3 min-h-0">
            <form onSubmit={agregarUnidad} className="flex gap-2 shrink-0">
              <input 
                type="text" 
                className="flex-1 bg-slate-50 dark:bg-slate-900 border border-slate-200 dark:border-slate-700 text-slate-800 dark:text-white px-3 py-2 rounded-xl font-bold text-xs uppercase outline-none focus:border-emerald-700 transition-colors placeholder-slate-400" 
                value={nuevaUnidad} 
                onChange={e => setNuevaUnidad(e.target.value)} 
                placeholder="Nueva Unidad..." 
                required 
              />
              <button type="submit" className="px-3.5 bg-emerald-600 hover:bg-emerald-700 text-white font-black rounded-xl text-base shadow-md transition-colors cursor-pointer">
                ＋
              </button>
            </form>

            <div className="flex-1 overflow-y-auto pr-1 space-y-1.5 custom-scrollbar min-h-0">
              {unidades.length === 0 ? (
                <p className="text-center text-[10px] font-bold text-slate-400 dark:text-slate-500 italic mt-4">No hay unidades registradas.</p>
              ) : (
                unidades.map(u => (
                  <div key={u.id} className="px-3 py-2 flex justify-between items-center text-[11px] font-black uppercase text-slate-700 dark:text-slate-200 bg-slate-50 dark:bg-slate-900/60 border border-slate-100 dark:border-slate-700 rounded-xl hover:border-emerald-200 transition-colors">
                    <span>{u.nombre_unidad}</span>
                    <button 
                      onClick={() => eliminarElemento('config_unidades', u.id, 'nombre_unidad', u.nombre_unidad)} 
                      className="p-1 bg-red-100 dark:bg-red-950/60 text-red-700 dark:text-red-400 hover:bg-red-600 hover:text-white rounded-lg transition-colors cursor-pointer text-xs"
                      title="Eliminar Unidad"
                    >
                      🗑️
                    </button>
                  </div>
                ))
              )}
            </div>
          </div>
        </div>

      </div>
    </div>
  );
}