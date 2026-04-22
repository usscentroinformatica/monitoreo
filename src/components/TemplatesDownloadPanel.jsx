import React, { useMemo } from 'react';

const PERIOD_START = new Date(2026, 3, 6, 0, 0, 0, 0); // 06/04/2026
const PERIOD_END = new Date(2026, 4, 30, 23, 59, 59, 999); // 30/05/2026

const DATE_KEYS = ['Fecha', 'FECHA', 'DIA', 'Dia', 'Columna 13', 'COLUMNA 13'];
const EFFICIENCY_KEYS = ['EFICIENCIA', 'Eficiencia', 'INDICE EFICIENCIA', 'Índice de Eficiencia'];
const EFFICACY_KEYS = ['EFICACIA', 'Eficacia', 'INDICE EFICACIA', 'Índice de Eficacia'];
const START_KEYS = ['inicio', 'INICIO', 'Hora Inicio', 'HORA INICIO'];
const END_KEYS = ['fin', 'FIN', 'Hora Fin', 'HORA FIN'];
const ZOOM_END_KEYS = ['FINALIZA LA CLASE (ZOOM)', 'Finaliza la Clase (Zoom)', 'Hora Finalización Zoom'];
const SESSIONS_PER_GROUP = 16;

const toDate = (value) => {
  const s = String(value || '').trim();
  if (!s) return null;

  const m1 = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
  if (m1) {
    const day = parseInt(m1[1], 10);
    const month = parseInt(m1[2], 10);
    const year = parseInt(m1[3].length === 2 ? `20${m1[3]}` : m1[3], 10);
    return new Date(year, month - 1, day, 12, 0, 0, 0);
  }

  const iso = new Date(s);
  return Number.isNaN(iso.getTime()) ? null : iso;
};

const toEfficiencyNumber = (value) => {
  const s = String(value || '').trim();
  if (!s) return NaN;
  const normalized = s.replace('%', '').replace(',', '.').trim();
  const num = parseFloat(normalized);
  return Number.isFinite(num) ? num : NaN;
};

const firstValue = (row, keys) => keys.map(k => row?.[k]).find(v => String(v || '').trim() !== '');

const parseSession = (value) => {
  const n = parseInt(String(value || '').trim(), 10);
  return Number.isFinite(n) ? n : 0;
};

const isRowMonitored = (row) => {
  const eff = toEfficiencyNumber(firstValue(row, EFFICIENCY_KEYS));
  if (Number.isFinite(eff)) return true;

  const hasStart = String(firstValue(row, START_KEYS) || '').trim() !== '';
  const hasEnd = String(firstValue(row, END_KEYS) || '').trim() !== '';
  const hasZoom = String(firstValue(row, ZOOM_END_KEYS) || '').trim() !== '';
  return (hasStart && hasEnd) || hasZoom;
};

function TemplatesDownloadPanel({ backupHistory = [] }) {
  const metrics = useMemo(() => {
    const latestBackupWithData = [...backupHistory]
      .sort((a, b) => new Date(b.date || 0).getTime() - new Date(a.date || 0).getTime())
      .find(b => Array.isArray(b.data) && b.data.length > 0);

    const rows = latestBackupWithData?.data || [];

    const today = new Date();
    const cutoff = today < PERIOD_END ? today : PERIOD_END;
    const periodRows = rows.filter((row) => {
      const rawDate = firstValue(row, DATE_KEYS);
      const d = toDate(rawDate);
      if (!d) return false;
      return d >= PERIOD_START && d <= cutoff;
    });

    const groupKeys = new Set(
      periodRows
        .map((r) => {
          const docente = String(r?.DOCENTE || '').trim();
          const curso = String(r?.CURSO || '').trim();
          const seccion = String(r?.SECCION || r?.['SECCIÓN'] || '').trim();
          if (!docente || !curso || !seccion) return '';
          return `${docente}|||${curso}|||${seccion}`;
        })
        .filter(Boolean)
    );

    const expectedSessions = groupKeys.size * SESSIONS_PER_GROUP;

    const monitoredSessionKeys = new Set(
      periodRows
        .filter((r) => {
          const s = parseSession(r?.SESION);
          return s >= 1 && s <= SESSIONS_PER_GROUP && isRowMonitored(r);
        })
        .map((r) => {
          const docente = String(r?.DOCENTE || '').trim();
          const curso = String(r?.CURSO || '').trim();
          const seccion = String(r?.SECCION || r?.['SECCIÓN'] || '').trim();
          const sesion = parseSession(r?.SESION);
          return `${docente}|||${curso}|||${seccion}|||${sesion}`;
        })
    );

    const monitoredSessions = monitoredSessionKeys.size;
    const avancePct = expectedSessions > 0 ? (monitoredSessions / expectedSessions) * 100 : 0;

    const docentes = new Set(
      periodRows
        .map(r => String(r?.DOCENTE || '').trim())
        .filter(Boolean)
    );

    const efficiencies = periodRows
      .map(r => toEfficiencyNumber(firstValue(r, EFFICIENCY_KEYS)))
      .filter(Number.isFinite);

    const avgEfficiency = efficiencies.length > 0
      ? efficiencies.reduce((sum, v) => sum + v, 0) / efficiencies.length
      : 0;

    const efficacies = periodRows
      .map(r => toEfficiencyNumber(firstValue(r, EFFICACY_KEYS)))
      .filter(Number.isFinite);

    const avgEfficacy = efficacies.length > 0
      ? efficacies.reduce((sum, v) => sum + v, 0) / efficacies.length
      : 0;

    const byDocente = new Map();
    periodRows.forEach((row) => {
      const docente = String(row?.DOCENTE || '').trim();
      if (!docente) return;
      const eff = toEfficiencyNumber(firstValue(row, EFFICIENCY_KEYS));
      if (!Number.isFinite(eff)) return;
      if (!byDocente.has(docente)) byDocente.set(docente, []);
      byDocente.get(docente).push(eff);
    });

    let topDocente = '';
    let topValue = -1;
    byDocente.forEach((values, docente) => {
      const avg = values.reduce((sum, v) => sum + v, 0) / values.length;
      if (avg > topValue) {
        topValue = avg;
        topDocente = docente;
      }
    });

    return {
      rowsSource: rows.length,
      rowsInPeriod: periodRows.length,
      monitoredSessions,
      expectedSessions,
      groupsCount: groupKeys.size,
      avancePct,
      docentesCount: docentes.size,
      avgEfficiency,
      avgEfficacy,
      topDocente,
      topValue,
      hasData: rows.length > 0
    };
  }, [backupHistory]);

  const safeProgress = Math.max(0, Math.min(100, metrics.avancePct));

  return (
    <div className="bg-white rounded-b-xl shadow-2xl p-8 lg:p-10 border border-[#d7f3fa]">
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-8">
        <section className="bg-[#f7fbff] rounded-xl border border-[#d7f3fa] p-6 flex flex-col items-center justify-center text-center">
          <h2 className="text-xs font-black tracking-[0.25em] text-[#5a2290] uppercase mb-4">Indicadores</h2>

          {/* Porcentaje principal */}
          <p className="text-8xl font-black leading-none" style={{ color: safeProgress >= 80 ? '#63a80f' : safeProgress >= 50 ? '#11acd3' : '#5a2290' }}>
            {safeProgress.toFixed(1)}%
          </p>
          <p className="text-sm text-slate-500 mt-2 mb-5">Avance al {new Date().toLocaleDateString('es-PE', { day: '2-digit', month: 'long', year: 'numeric' })}</p>

          {/* Barra de progreso */}
          <div className="w-full h-4 bg-[#e8f6fb] rounded-full overflow-hidden mb-6">
            <div
              className="h-4 rounded-full transition-all duration-500"
              style={{ width: `${safeProgress}%`, background: safeProgress >= 80 ? '#63ed12' : safeProgress >= 50 ? '#11acd3' : '#5a2290' }}
            />
          </div>

          {/* Métricas secundarias */}
          <div className="grid grid-cols-4 gap-2 w-full text-sm">
            <div className="bg-white rounded-lg border border-[#e8f6fb] p-3">
              <p className="text-slate-400 text-xs">Docentes</p>
              <p className="text-xl font-black text-[#5a2290]">{metrics.docentesCount}</p>
            </div>
            <div className="bg-white rounded-lg border border-[#e8f6fb] p-3">
              <p className="text-slate-400 text-xs">Sesiones</p>
              <p className="text-xl font-black text-[#11acd3]">{metrics.monitoredSessions}<span className="text-xs text-slate-400">/{metrics.expectedSessions}</span></p>
            </div>
            <div className="bg-white rounded-lg border border-[#e8f6fb] p-3">
              <p className="text-slate-400 text-xs">Eficiencia</p>
              <p className="text-xl font-black text-[#63ed12]">{metrics.avgEfficiency.toFixed(1)}%</p>
            </div>
            <div className="bg-white rounded-lg border border-[#e8f6fb] p-3">
              <p className="text-slate-400 text-xs">Eficacia</p>
              <p className="text-xl font-black text-[#11acd3]">{metrics.avgEfficacy.toFixed(1)}%</p>
            </div>
          </div>

          {!metrics.hasData && (
            <p className="text-xs text-[#5a2290] font-semibold mt-4">Sin datos guardados aún.</p>
          )}
        </section>

        <section className="bg-white rounded-xl border border-[#d7f3fa] p-6 text-center">
          <h3 className="text-xl font-extrabold text-[#5a2290] mb-2">Descarga las plantillas</h3>
          <p className="text-sm text-slate-600 mb-6">Usa estos archivos oficiales para iniciar el monitoreo.</p>

          <div className="flex flex-col items-center gap-3">
            <a
              href="/guia monitoreo.pdf"
              target="_blank"
              rel="noopener noreferrer"
              className="px-4 py-2 bg-[#5a2290] text-white rounded hover:bg-[#11acd3] font-bold shadow w-72 text-center"
            >
              📘 Ver Guía de Monitoreo (PDF)
            </a>

            <a
              href="/CINF%20202601%20ABRIL.xlsx"
              download="CINF 202601 ABRIL.xlsx"
              className="px-4 py-2 bg-[#63ed12] text-[#124007] rounded hover:bg-[#11acd3] hover:text-white font-bold shadow w-72 text-center"
            >
              📥 Descargar plantilla Excel
            </a>

            <a
              href="/meetings_Docentes_CIS_2026_04_06_2026_04_12.csv"
              download="meetings_Docentes_CIS_2026_04_06_2026_04_12.csv"
              className="px-4 py-2 bg-[#11acd3] text-white rounded hover:bg-[#5a2290] font-bold shadow w-72 text-center"
            >
              📊 Descargar reporte Zoom
            </a>
          </div>
        </section>
      </div>
    </div>
  );
}

export default TemplatesDownloadPanel;
