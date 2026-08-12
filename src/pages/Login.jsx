import { useState } from 'react';
import { supabase } from '../lib/supabase';

export default function Login({ setSession }) {
  const [email, setEmail] = useState('');
  const [password, setPassword] = useState('');
  const [loading, setLoading] = useState(false);
  const [errorMsg, setErrorMsg] = useState(null);
  const [rolSeleccionado, setRolSeleccionado] = useState(null); // 'admin' u 'operario'

  const handleLogin = async (e) => {
    e.preventDefault();
    setLoading(true);
    setErrorMsg(null);
    
    const { data, error } = await supabase.auth.signInWithPassword({
      email,
      password,
    });

    if (error) {
      if (error.message === 'Invalid login credentials') {
        setErrorMsg('El correo o la contraseña son incorrectos. Verifique e intente de nuevo.');
      } else {
        setErrorMsg(error.message);
      }
    } else {
      setSession(data.session);
    }
    setLoading(false);
  };

  return (
    <div className="min-h-screen flex flex-col lg:flex-row bg-[#eef6f0] font-sans overflow-hidden">
      
      {/* 🌿 PANEL IZQUIERDO: VERDE UN POCO MÁS CLARO Y FRESCO */}
      <div className="lg:w-4/12 xl:w-3/12 bg-gradient-to-br from-green-900 via-green-800 to-emerald-900 p-8 flex flex-col justify-between text-white relative z-10 min-h-screen shadow-2xl">
        
        {/* Título GRANJA WP Grande y Centrado */}
        <div className="text-center pt-4">
          <h1 className="text-2xl font-black tracking-widest uppercase text-white drop-shadow-md">
            GRANJA WP
          </h1>
        </div>

        {/* 🖼️ ILUSTRACIÓN GRANDE CENTRADA */}
        <div className="my-auto flex flex-col items-center justify-center py-6">
          <div className="bg-green-950/40 p-6 rounded-3xl border border-green-600/40 shadow-2xl backdrop-blur-sm max-w-[200px] w-full flex items-center justify-center">
            <img 
              src="/Granja.png" 
              alt="Granja WP Ilustración" 
              className="w-full h-auto max-h-[130px] object-contain rounded-2xl drop-shadow-xl"
            />
          </div>
        </div>

        {/* Bloque de Bienvenida Centrado */}
        <div className="space-y-3 text-center pb-6">
          <div className="inline-block bg-green-700/80 border border-green-500/40 px-3 py-1 rounded-full text-[9px] font-black uppercase tracking-widest text-green-100 shadow-sm">
            Plataforma Agrícola
          </div>
          <h2 className="text-xl lg:text-2xl font-black tracking-tight leading-tight text-white">
            Bienvenido a su centro de control.
          </h2>
          <p className="text-green-100/90 text-[11px] font-medium leading-relaxed">
            Administre sus Invernaderos, cosechas, despachos y control de egresos de manera eficiente.
          </p>
        </div>

        <div className="text-[10px] text-green-200/70 font-bold uppercase tracking-wider text-center pb-2">
          © {new Date().getFullYear()} Intepe
        </div>
      </div>

      {/* 🔐 PANEL DERECHO: TARJETAS Y FORMULARIO SOBRE EL FONDO VERDE MANZANA */}
      <div className="lg:w-8/12 xl:w-9/12 flex items-center justify-center p-6 lg:p-12 relative z-20 bg-[#eef6f0]">
        <div className="w-full max-w-md space-y-6">
          
          {!rolSeleccionado ? (
            /* 🎴 PASO 1: TARJETAS DE SELECCIÓN DE ROL */
            <div className="space-y-4 animate-in fade-in slide-in-from-right-4 duration-300">
              <div className="text-center mb-6">
                <h3 className="text-2xl font-black text-slate-900 tracking-tight">¿Cómo deseas ingresar?</h3>
                <p className="text-xs font-bold text-gray-500 uppercase tracking-widest mt-1">Seleccione su perfil de usuario</p>
              </div>

              {/* Tarjeta Administrador */}
              <div 
                onClick={() => setRolSeleccionado('admin')}
                className="bg-white p-5 rounded-2xl shadow-sm border border-green-200/60 hover:border-green-600 hover:shadow-md transition-all cursor-pointer flex items-center justify-between group"
              >
                <div className="flex items-center gap-4">
                  <div className="w-12 h-12 rounded-xl bg-amber-50 text-amber-600 flex items-center justify-center text-xl font-black shadow-inner">
                    👑
                  </div>
                  <div>
                    <h4 className="font-black text-slate-800 text-sm uppercase">Administrador</h4>
                    <p className="text-[11px] font-bold text-slate-400">Control total y financiero</p>
                  </div>
                </div>
                <span className="text-slate-300 group-hover:text-green-700 font-black text-lg transition-colors">›</span>
              </div>

              {/* Tarjeta Operador */}
              <div 
                onClick={() => setRolSeleccionado('operario')}
                className="bg-white p-5 rounded-2xl shadow-sm border border-green-200/60 hover:border-green-600 hover:shadow-md transition-all cursor-pointer flex items-center justify-between group"
              >
                <div className="flex items-center gap-4">
                  <div className="w-12 h-12 rounded-xl bg-green-50 text-green-700 flex items-center justify-center text-xl font-black shadow-inner">
                    👨‍🌾
                  </div>
                  <div>
                    <h4 className="font-black text-slate-800 text-sm uppercase">Operador</h4>
                    <p className="text-[11px] font-bold text-slate-400">Registro de cosecha e inventario</p>
                  </div>
                </div>
                <span className="text-slate-300 group-hover:text-green-700 font-black text-lg transition-colors">›</span>
              </div>
            </div>
          ) : (
            /* 📝 PASO 2: FORMULARIO DE ACCESO SEGÚN EL ROL */
            <div className="bg-white p-8 rounded-3xl shadow-xl border border-green-100 space-y-6 animate-in fade-in slide-in-from-left-4 duration-300 relative">
              
              <button 
                type="button"
                onClick={() => { setRolSeleccionado(null); setErrorMsg(null); }}
                className="text-xs font-black text-green-700 hover:underline flex items-center gap-1 uppercase tracking-wider cursor-pointer"
              >
                ← Volver a selección
              </button>

              <div className="space-y-1">
                <div className="flex items-center gap-2">
                  <span className="text-lg">{rolSeleccionado === 'admin' ? '👑' : '👨‍🌾'}</span>
                  <h3 className="text-xl font-black text-slate-900 uppercase tracking-tight">
                    Acceso {rolSeleccionado === 'admin' ? 'Administrador' : 'Operador'}
                  </h3>
                </div>
                <p className="text-[10px] font-bold text-gray-400 uppercase tracking-widest">Ingrese sus credenciales de acceso</p>
              </div>

              <form onSubmit={handleLogin} className="space-y-4">
                
                {errorMsg && (
                  <div className="p-3.5 bg-red-50 border border-red-200 rounded-2xl flex items-start gap-2.5 animate-in fade-in">
                    <span className="text-lg">⚠️</span>
                    <div className="text-left">
                      <p className="text-[9px] font-black text-red-500 uppercase tracking-widest leading-none">Error</p>
                      <p className="text-xs font-bold text-red-800 mt-0.5 leading-tight">{errorMsg}</p>
                    </div>
                  </div>
                )}

                <div className="space-y-1">
                  <label className="text-[10px] font-black text-slate-400 uppercase tracking-wider px-1">Correo Electrónico</label>
                  <input 
                    type="email" 
                    placeholder="ejemplo@granja.com" 
                    className="w-full border-2 border-slate-100 p-3.5 rounded-2xl outline-none focus:border-green-700 font-bold text-sm bg-slate-50/50 transition-all text-slate-800"
                    value={email}
                    onChange={(e) => setEmail(e.target.value)}
                    required
                  />
                </div>

                <div className="space-y-1">
                  <label className="text-[10px] font-black text-slate-400 uppercase tracking-wider px-1">Contraseña</label>
                  <input 
                    type="password" 
                    placeholder="••••••••" 
                    className="w-full border-2 border-slate-100 p-3.5 rounded-2xl outline-none focus:border-green-700 font-bold text-sm bg-slate-50/50 transition-all text-slate-800"
                    value={password}
                    onChange={(e) => setPassword(e.target.value)}
                    required
                  />
                </div>

                <button 
                  disabled={loading}
                  className="w-full bg-green-700 text-white font-black py-4 rounded-2xl uppercase tracking-widest text-xs hover:bg-green-800 transition-all shadow-lg shadow-green-700/20 disabled:opacity-50 cursor-pointer mt-2"
                >
                  {loading ? 'Verificando...' : 'Entrar al Sistema'}
                </button>
              </form>
            </div>
          )}

        </div>
      </div>

    </div>
  );
}