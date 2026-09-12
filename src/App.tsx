import { useState } from 'react';
import { BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer } from 'recharts';
import { UploadCloud, CheckCircle2, AlertTriangle, AlertCircle, Activity } from 'lucide-react';
import * as XLSX from 'xlsx';

const kpisDataInicial = [
  { indicador: 'OEE1', actual: 0, meta: 0 },
  { indicador: 'OEE2', actual: 0, meta: 0 },
  { indicador: 'Rendimiento', actual: 0, meta: 0 },
  { indicador: 'Efectividad', actual: 0, meta: 0 },
];

export default function App() {
  const [totalProducido, setTotalProducido] = useState<number>(0);
  const [datosGrafico, setDatosGrafico] = useState<any[]>(kpisDataInicial);
  const [datosResumen, setDatosResumen] = useState<any[]>([]);
  const [datosTablaPT, setDatosTablaPT] = useState<any[]>([
    { linea: '-', codigo: '-', descripcion: 'Esperando Excel...', kilos: 0, estado: 'No Evaluado' }
  ]);

  const manejarSubidaExcel = (e: React.ChangeEvent<HTMLInputElement>) => {
    const archivo = e.target.files?.[0];
    if (!archivo) return;

    const lector = new FileReader();
    lector.onload = (evento) => {
      try {
        const buffer = evento.target?.result;
        if (buffer instanceof ArrayBuffer) {
          const workbook = XLSX.read(buffer, { type: 'array' });
          
          // =================================================================
          // 1. PROCESAR HOJA "RESUMEN"
          // =================================================================
          const nombreHojaResumen = workbook.SheetNames.find(n => n.toUpperCase().includes('RESUMEN'));
          if (nombreHojaResumen) {
            const hojaResumen = workbook.Sheets[nombreHojaResumen];
            const matrizResumen = XLSX.utils.sheet_to_json<any[]>(hojaResumen, { header: 1, raw: false });
            
            let headerRow = -1;
            let colKPI = -1, colOEE1 = -1, colOEE2 = -1, colRend = -1, colEfect = -1, colWaste = -1;

            for (let i = 0; i < matrizResumen.length; i++) {
              const fila = matrizResumen[i];
              if (!fila) continue;
              for (let j = 0; j < fila.length; j++) {
                const celda = String(fila[j] || '').toLowerCase().trim();
                if (celda === 'kpi') { headerRow = i; colKPI = j; }
                if (celda.includes('oee1')) colOEE1 = j;
                if (celda.includes('oee2')) colOEE2 = j;
                if (celda.includes('rend.')) colRend = j;
                if (celda.includes('efect')) colEfect = j;
                if (celda.includes('waste') || celda.includes('desperdicio')) colWaste = j;
              }
              if (headerRow !== -1) break;
            }

            if (headerRow !== -1) {
              const resumenTemp = [];
              let rowMeta: any = null;
              let rowResultado: any = null;

              const aNumero = (val: any) => {
                if (!val) return 0;
                const num = Number(String(val).replace(/[%,\s]/g, ''));
                return isNaN(num) ? 0 : num;
              };

              for (let i = headerRow + 1; i < matrizResumen.length; i++) {
                const fila = matrizResumen[i];
                if (!fila || !fila[colKPI]) continue;
                
                const nombreKpi = String(fila[colKPI]).trim();
                if (nombreKpi === '') continue;

                const obj = {
                  kpi: nombreKpi,
                  oee1: aNumero(fila[colOEE1]),
                  oee2: aNumero(fila[colOEE2]),
                  rendimiento: aNumero(fila[colRend]),
                  efectividad: aNumero(fila[colEfect]),
                  waste: aNumero(fila[colWaste])
                };
                resumenTemp.push(obj);

                if (nombreKpi.toUpperCase() === 'META%' && !rowMeta) rowMeta = obj;
                if (nombreKpi.toUpperCase().includes('RESULTADO GBC') || nombreKpi.toUpperCase().includes('RESULTADO')) rowResultado = obj;
              }

              setDatosResumen(resumenTemp);

              if (rowMeta && rowResultado) {
                setDatosGrafico([
                  { indicador: 'OEE1', actual: rowResultado.oee1, meta: rowMeta.oee1 },
                  { indicador: 'OEE2', actual: rowResultado.oee2, meta: rowMeta.oee2 },
                  { indicador: 'Rendimiento', actual: rowResultado.rendimiento, meta: rowMeta.rendimiento },
                  { indicador: 'Efectividad', actual: rowResultado.efectividad, meta: rowMeta.efectividad },
                ]);
              }
            }
          }

          // =================================================================
          // 2. PROCESAR HOJA "PT" (Detalle de Producción)
          // =================================================================
          const nombreHojaPT = workbook.SheetNames.find(n => n.toUpperCase() === 'PT') || workbook.SheetNames[0];
          const hojaPT = workbook.Sheets[nombreHojaPT];
          const matrizPT = XLSX.utils.sheet_to_json<any[]>(hojaPT, { header: 1 });
          
          let idxLinea = -1, idxCodigo = -1, idxDesc = -1, idxKilos = -1;
          let sumaKilos = 0;
          const nuevasFilasPT: any[] = [];

          for (let i = 0; i < matrizPT.length; i++) {
            const fila = matrizPT[i];
            if (!fila || fila.length === 0) continue;

            if (idxCodigo === -1) {
              for (let j = 0; j < fila.length; j++) {
                const celda = String(fila[j] || '').toLowerCase().trim();
                if (celda.includes('línea') || celda.includes('linea')) idxLinea = j;
                if (celda.includes('código') || celda.includes('codigo') || celda === 'cod') idxCodigo = j;
                if (celda.includes('descripción') || celda.includes('descripcion') || celda.includes('producto')) idxDesc = j;
                // Buscamos tanto kilos de desperdicio como columnas de producción o valor
                if (celda.includes('desperdicio') || celda.includes('masas') || celda.includes('kilos') || celda.includes('kg') || celda.includes('válido')) idxKilos = j;
              }
            } else {
              const codigoFila = fila[idxCodigo];
              if (codigoFila && String(codigoFila).trim() !== '' && String(codigoFila).toLowerCase() !== 'total') {
                const kilosValidos = idxKilos !== -1 ? Number(fila[idxKilos]) || 0 : 0;
                sumaKilos += kilosValidos;

                nuevasFilasPT.push({
                  linea: idxLinea !== -1 ? (fila[idxLinea] || 'N/A') : 'N/A',
                  codigo: codigoFila,
                  descripcion: idxDesc !== -1 ? (fila[idxDesc] || 'N/A') : 'N/A',
                  kilos: kilosValidos,
                  estado: kilosValidos === 0 ? 'Excelente' : (kilosValidos <= 20 ? 'Aceptable' : 'Alerta')
                });
              }
            }
          }

          if (nuevasFilasPT.length > 0) {
              setDatosTablaPT(nuevasFilasPT);
              setTotalProducido(sumaKilos);
          }
          
          alert("¡Reporte Multi-Hoja Procesado Correctamente!");
        }
      } catch (error) {
        console.error("Error procesando Excel:", error);
      }
    };
    lector.readAsArrayBuffer(archivo);
  };

  const renderEstado = (estado: string) => {
    switch (estado) {
      case 'Excelente':
        return <span className="flex items-center justify-center gap-1 text-green-700 bg-green-100 px-2 py-1 rounded text-xs font-bold"><CheckCircle2 size={14} /> 0 DEFECTO</span>;
      case 'Aceptable':
        return <span className="flex items-center justify-center gap-1 text-yellow-700 bg-yellow-100 px-2 py-1 rounded text-xs font-bold"><AlertTriangle size={14} /> EN LÍMITE</span>;
      default:
        return <span className="flex items-center justify-center gap-1 text-red-600 bg-red-50 px-2 py-1 rounded text-xs font-bold"><AlertCircle size={14} /> ALERTA</span>;
    }
  };

  const oee2Actual = datosGrafico.find(d => d.indicador === 'OEE2')?.actual || 0;
  const wasteActual = datosResumen.find(d => d.kpi.toUpperCase().includes('RESULTADO'))?.waste || 0;

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
        
        {/* TARJETAS DE SUPERFICIE */}
        <div className="grid grid-cols-1 md:grid-cols-4 gap-4 mb-6">
          <div className="bg-white rounded-xl shadow-sm p-5 border border-gray-100">
            <h3 className="text-sm font-semibold text-gray-500 mb-1">OEE2 General (Resultado)</h3>
            <div className="flex items-end gap-2"><span className="text-3xl font-bold">{oee2Actual}%</span></div>
          </div>
          <div className="bg-white rounded-xl shadow-sm p-5 border border-gray-100">
            <h3 className="text-sm font-semibold text-gray-500 mb-1">% Total Waste</h3>
            <div className="flex items-end gap-2">
              <span className={`text-3xl font-bold ${wasteActual > 1.5 ? 'text-red-600' : 'text-green-600'}`}>
                {wasteActual}%
              </span>
            </div>
          </div>
          <div className="bg-white rounded-xl shadow-sm p-5 border border-gray-100">
            <h3 className="text-sm font-semibold text-gray-500 mb-1">Volumen Consolidado</h3>
            <div className="flex items-end gap-2 text-blue-800"><span className="text-3xl font-bold">{totalProducido.toLocaleString(undefined, { maximumFractionDigits: 2 })}</span></div>
          </div>
          <div className="bg-white rounded-xl shadow-sm p-5 border border-gray-100 flex items-center justify-between">
             <div>
              <h3 className="text-sm font-semibold text-gray-500 mb-1">Estado General</h3>
              <span className={`text-lg font-bold ${oee2Actual >= 92 ? 'text-green-600' : 'text-red-600'}`}>
                {oee2Actual >= 92 ? 'En Meta' : 'Bajo Meta'}
              </span>
            </div>
            <div className={`w-10 h-10 rounded-full shadow-[0_0_15px_rgba(34,197,94,0.5)] border-4 animate-pulse ${oee2Actual >= 92 ? 'bg-green-500 border-green-100' : 'bg-red-500 border-red-100 shadow-[0_0_15px_rgba(239,68,68,0.5)]'}`}></div>
          </div>
        </div>

        {/* GRÁFICO */}
        <div className="bg-white rounded-xl shadow-sm border border-gray-100 p-6 mb-6">
          <h2 className="text-lg font-bold text-gray-800 mb-4 flex items-center gap-2"><Activity size={20} className="text-blue-800"/> Cumplimiento de Metas (Hoja: RESUMEN)</h2>
          <div className="h-72 w-full">
            <ResponsiveContainer width="100%" height="100%">
              <BarChart data={datosGrafico} margin={{ top: 5, right: 30, left: 0, bottom: 5 }}>
                <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e5e7eb" />
                <XAxis dataKey="indicador" axisLine={false} tickLine={false} />
                <YAxis axisLine={false} tickLine={false} tickFormatter={(value) => value + "%"} />
                <Tooltip cursor={{fill: '#f3f4f6'}} />
                <Legend />
                <Bar dataKey="actual" name="Resultado Actual %" fill="#0033A0" radius={[4, 4, 0, 0]} />
                <Bar dataKey="meta" name="Meta %" fill="#9ca3af" radius={[4, 4, 0, 0]} />
              </BarChart>
            </ResponsiveContainer>
          </div>
        </div>

        {/* TABLAS INFERIORES */}
        <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
          
          <div className="bg-white rounded-xl shadow-sm border border-gray-100 overflow-hidden">
            <div className="p-5 border-b border-gray-100 bg-gray-50 flex justify-between items-center">
              <h2 className="text-md font-bold text-gray-800">Tabla de Metas (Hoja: RESUMEN)</h2>
            </div>
            <div className="max-h-96 overflow-y-auto">
              <table className="w-full text-left border-collapse">
                <thead className="sticky top-0 bg-white shadow-sm z-10">
                  <tr className="text-gray-500 text-[10px] uppercase tracking-wider border-b border-gray-100">
                    <th className="p-3 font-semibold">KPI / Línea</th>
                    <th className="p-3 font-semibold text-center">OEE1</th>
                    <th className="p-3 font-semibold text-center">OEE2</th>
                    <th className="p-3 font-semibold text-center">Rendimiento</th>
                    <th className="p-3 font-semibold text-center">% Waste</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-gray-100">
                  {datosResumen.length === 0 && (
                    <tr><td colSpan={5} className="p-4 text-center text-sm text-gray-400">Sube el reporte...</td></tr>
                  )}
                  {datosResumen.map((fila, index) => (
                    <tr key={index} className="hover:bg-blue-50/50 transition-colors">
                      <td className="p-3 text-xs font-bold text-gray-700">{fila.kpi}</td>
                      <td className="p-3 text-xs text-center">{fila.oee1}%</td>
                      <td className="p-3 text-xs font-medium text-blue-800 text-center">{fila.oee2}%</td>
                      <td className="p-3 text-xs text-center">{fila.rendimiento}%</td>
                      <td className="p-3 text-xs text-center text-red-600">{fila.waste}%</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>

          <div className="bg-white rounded-xl shadow-sm border border-gray-100 overflow-hidden">
            <div className="p-5 border-b border-gray-100 bg-gray-50 flex justify-between items-center">
              <h2 className="text-md font-bold text-gray-800">Detalle de Producción (Hoja: PT)</h2>
              <span className="text-xs font-semibold text-blue-800 bg-blue-100 px-2 py-1 rounded-full">{datosTablaPT.length} regs</span>
            </div>
            <div className="max-h-96 overflow-y-auto">
              <table className="w-full text-left border-collapse">
                <thead className="sticky top-0 bg-white shadow-sm z-10">
                  <tr className="text-gray-500 text-[10px] uppercase tracking-wider border-b border-gray-100">
                    <th className="p-3 font-semibold">Cód.</th>
                    <th className="p-3 font-semibold">Descripción</th>
                    <th className="p-3 font-semibold text-right">Valor / Kilos</th>
                    <th className="p-3 font-semibold text-center">Estado</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-gray-100">
                  {datosTablaPT.map((fila, index) => (
                    <tr key={index} className="hover:bg-blue-50/50 transition-colors">
                      <td className="p-3 text-xs font-medium text-gray-900">{fila.codigo}</td>
                      <td className="p-3 text-xs text-gray-600 truncate max-w-[150px]">{fila.descripcion}</td>
                      <td className="p-3 text-xs font-medium text-gray-900 text-right">{fila.kilos.toLocaleString()}</td>
                      <td className="p-3 flex justify-center">{renderEstado(fila.estado)}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>

        </div>
      </main>
    </div>
  );
}