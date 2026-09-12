import { useState } from 'react';
import { BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer } from 'recharts';
import { UploadCloud, CheckCircle2, AlertTriangle, AlertCircle } from 'lucide-react';
import * as XLSX from 'xlsx';

const kpisData = [
  { indicador: 'OEE1', actual: 33.2, meta: 32.6 },
  { indicador: 'OEE2', actual: 93.8, meta: 92.0 },
  { indicador: 'Rendimiento', actual: 97.1, meta: 97.5 },
  { indicador: 'Efectividad', actual: 96.5, meta: 86.0 },
];

export default function App() {
  const [totalProducido, setTotalProducido] = useState<number>(0);
  const [datosTabla, setDatosTabla] = useState<any[]>([
    { linea: '-', codigo: '-', descripcion: 'Sube tu Excel para extraer los datos...', kilos: 0, estado: 'No Evaluado' }
  ]);

  // Función "inteligente" (Búsqueda Difusa) para encontrar columnas
  const encontrarValor = (fila: any, palabrasClave: string[]) => {
    if (!fila || typeof fila !== 'object') return null;
    const llaves = Object.keys(fila);
    const llaveEncontrada = llaves.find(llave => 
      palabrasClave.some(palabra => llave.toLowerCase().includes(palabra.toLowerCase()))
    );
    return llaveEncontrada ? fila[llaveEncontrada] : null;
  };

  const manejarSubidaExcel = (e: React.ChangeEvent<HTMLInputElement>) => {
    const archivo = e.target.files?.[0];
    if (!archivo) return;

    const lector = new FileReader();
    lector.onload = (evento) => {
      try {
        const buffer = evento.target?.result;
        if (buffer instanceof ArrayBuffer) {
          const workbook = XLSX.read(buffer, { type: 'array' });
          
          // 1. Apuntamos a la pestaña PT o a la primera que exista
          const nombreHoja = workbook.SheetNames.includes('PT') ? 'PT' : workbook.SheetNames[0];
          const hoja = workbook.Sheets[nombreHoja];
          
          // 2. range: 1 salta la fila 1 (los títulos combinados que rompen la lectura)
          const datosJson = XLSX.utils.sheet_to_json(hoja, { range: 1, defval: null });
          
          let sumaKilos = 0;
          const nuevasFilas: any[] = [];

          // 3. Procesamos fila por fila con nuestra lógica inteligente
          datosJson.forEach((fila: any) => {
            const kilosCrudos = encontrarValor(fila, ['masas', 'desperdicio', 'kilos', 'kg', 'peso', 'piezas', 'elaboradas']);
            const kilosFila = Number(kilosCrudos);
            
            // Filtramos basuras o filas vacías (solo procesamos si hay kilos válidos)
            if (!isNaN(kilosFila) && kilosFila > 0) {
               sumaKilos += kilosFila;
               
               const lineaFila = encontrarValor(fila, ['linea', 'línea', 'area', 'pt']);
               const codigoFila = encontrarValor(fila, ['codigo', 'código', 'cod']);
               const descripcionFila = encontrarValor(fila, ['descripción', 'descripcion', 'producto', 'nombre']);

               nuevasFilas.push({
                 linea: lineaFila || 'PAN BOLLERIA',
                 codigo: codigoFila || 'N/A',
                 descripcion: descripcionFila || 'N/A',
                 kilos: kilosFila,
                 // Lógica de semáforo básica
                 estado: kilosFila > 1000 ? 'Cumple' : (kilosFila === 0 ? 'No Evaluado' : 'Alerta')
               });
            }
          });

          if (nuevasFilas.length > 0) {
              setDatosTabla(nuevasFilas);
              setTotalProducido(sumaKilos);
              alert("¡Éxito! Se procesaron productos automáticamente usando extracción difusa.");
          } else {
              alert("No se encontraron productos. Verifica que la pestaña PT tenga datos.");
          }
        }
      } catch (error) {
        console.error("Error procesando Excel:", error);
        alert("Hubo un problema procesando el archivo.");
      }
    };
    lector.readAsArrayBuffer(archivo);
  };

  const renderEstado = (estado: string) => {
    switch (estado) {
      case 'Cumple':
        return <span className="flex items-center justify-center gap-1 text-green-600 bg-green-50 px-2 py-1 rounded text-xs font-bold"><CheckCircle2 size={14} /> CUMPLE</span>;
      case 'Alerta':
        return <span className="flex items-center justify-center gap-1 text-red-600 bg-red-50 px-2 py-1 rounded text-xs font-bold"><AlertCircle size={14} /> ALERTA</span>;
      default:
        return <span className="flex items-center justify-center gap-1 text-gray-500 bg-gray-100 px-2 py-1 rounded text-xs font-bold"><AlertTriangle size={14} /> N/E</span>;
    }
  };

  return (
    <div className="min-h-screen bg-gray-50 font-sans text-gray-800">
      <header className="bg-white shadow-sm border-b border-gray-200">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-4 flex justify-between items-center">
          <div className="flex items-center gap-3">
            <div className="w-8 h-8 bg-blue-800 rounded-md flex items-center justify-center font-bold text-white text-xs">BIM</div>
            <h1 className="text-2xl font-bold text-gray-800">Dashboard de Producción (Smart Scan)</h1>
          </div>
          <label className="cursor-pointer bg-blue-800 hover:bg-blue-700 text-white px-4 py-2 rounded-lg text-sm font-medium transition-colors flex items-center gap-2 shadow-sm">
            <UploadCloud size={18} />
            Escanear Excel (.xlsx)
            <input type="file" className="hidden" accept=".xlsx, .xls" onChange={manejarSubidaExcel} />
          </label>
        </div>
      </header>

      <main className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-8">
        <div className="grid grid-cols-1 md:grid-cols-4 gap-4 mb-6">
          <div className="bg-white rounded-xl shadow-sm p-5 border border-gray-100">
            <h3 className="text-sm font-semibold text-gray-500 mb-1">OEE General (Promedio)</h3>
            <div className="flex items-end gap-2"><span className="text-3xl font-bold">93.8%</span></div>
          </div>
          <div className="bg-white rounded-xl shadow-sm p-5 border border-gray-100">
            <h3 className="text-sm font-semibold text-gray-500 mb-1">Desperdicio Total</h3>
            <div className="flex items-end gap-2"><span className="text-3xl font-bold text-red-600">1.38%</span></div>
          </div>
          <div className="bg-white rounded-xl shadow-sm p-5 border border-gray-100">
            <h3 className="text-sm font-semibold text-gray-500 mb-1">Total Masas Detectadas</h3>
            <div className="flex items-end gap-2 text-blue-800"><span className="text-3xl font-bold">{totalProducido.toLocaleString(undefined, { maximumFractionDigits: 2 })} kg</span></div>
          </div>
          <div className="bg-white rounded-xl shadow-sm p-5 border border-gray-100 flex items-center justify-between">
             <div>
              <h3 className="text-sm font-semibold text-gray-500 mb-1">Estado General</h3>
              <span className="text-lg font-bold text-green-600">En Meta</span>
            </div>
            <div className="w-10 h-10 rounded-full bg-green-500 shadow-[0_0_15px_rgba(34,197,94,0.5)] border-4 border-green-100 animate-pulse"></div>
          </div>
        </div>

        <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
          <div className="lg:col-span-2 bg-white rounded-xl shadow-sm border border-gray-100 p-6">
            <h2 className="text-lg font-bold text-gray-800 mb-4">Cumplimiento de Metas (Smart KPIs)</h2>
            <div className="h-72 w-full">
              <ResponsiveContainer width="100%" height="100%">
                <BarChart data={kpisData} margin={{ top: 5, right: 30, left: 0, bottom: 5 }}>
                  <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e5e7eb" />
                  <XAxis dataKey="indicador" axisLine={false} tickLine={false} />
                  <YAxis axisLine={false} tickLine={false} tickFormatter={(value) => value + "%"} />
                  <Tooltip cursor={{fill: '#f3f4f6'}} />
                  <Legend />
                  <Bar dataKey="actual" name="Logro Actual %" fill="#0033A0" radius={[4, 4, 0, 0]} />
                  <Bar dataKey="meta" name="Meta %" fill="#9ca3af" radius={[4, 4, 0, 0]} />
                </BarChart>
              </ResponsiveContainer>
            </div>
          </div>

          <div className="bg-white rounded-xl shadow-sm border border-gray-100 p-6">
             <h2 className="text-lg font-bold text-gray-800 mb-4">Módulo de Alertas (Simulado)</h2>
             <div className="space-y-4">
                <div className="p-3 bg-red-50 rounded-lg border border-red-100">
                  <p className="text-sm font-bold text-red-800">Pan Artesano Blanco</p>
                  <p className="text-xs text-red-600 mt-1">Desviación detectada por encima del 1.5% permitido.</p>
                </div>
                <div className="p-3 bg-blue-50 rounded-lg border border-blue-100">
                  <p className="text-sm font-bold text-blue-800">Integración IA</p>
                  <p className="text-xs text-blue-700 mt-1">Extracción de datos basada en búsqueda de patrones semánticos activos.</p>
                </div>
             </div>
          </div>
        </div>

        <div className="mt-6 bg-white rounded-xl shadow-sm border border-gray-100 overflow-hidden">
          <div className="p-6 border-b border-gray-100 flex justify-between items-center bg-gray-50">
            <h2 className="text-lg font-bold text-gray-800">Detalle de Producción (Extraído Automáticamente)</h2>
            <span className="text-sm font-semibold text-blue-800 bg-blue-100 px-3 py-1 rounded-full">{datosTabla.length} filas mapeadas</span>
          </div>
          
          <div className="max-h-96 overflow-y-auto">
            <table className="w-full text-left border-collapse">
              <thead className="sticky top-0 bg-white shadow-sm z-10">
                <tr className="text-gray-500 text-xs uppercase tracking-wider border-b border-gray-100">
                  <th className="p-4 font-semibold">Línea</th>
                  <th className="p-4 font-semibold">Cód. Producto</th>
                  <th className="p-4 font-semibold">Descripción</th>
                  <th className="p-4 font-semibold text-right">Volumen (kg/piezas)</th>
                  <th className="p-4 font-semibold text-center">Evaluación Automática</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-gray-100">
                {datosTabla.map((fila, index) => (
                  <tr key={index} className="hover:bg-blue-50/50 transition-colors">
                    <td className="p-4 text-sm text-gray-600">{fila.linea}</td>
                    <td className="p-4 text-sm font-medium text-gray-900">{fila.codigo}</td>
                    <td className="p-4 text-sm text-gray-600">{fila.descripcion}</td>
                    <td className="p-4 text-sm font-medium text-gray-900 text-right">{fila.kilos.toLocaleString(undefined, { maximumFractionDigits: 2 })}</td>
                    <td className="p-4 flex justify-center">{renderEstado(fila.estado)}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      </main>
    </div>
  );
}