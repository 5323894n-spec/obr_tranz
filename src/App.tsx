import { useState } from 'react';
import { Button } from '@/components/ui/button';
import { Card, CardContent, CardDescription, CardFooter, CardHeader, CardTitle } from '@/components/ui/card';
import { Input } from '@/components/ui/input';
import { Label } from '@/components/ui/label';
import { Alert, AlertDescription, AlertTitle } from '@/components/ui/alert';
import { Progress } from '@/components/ui/progress';
import { FileCheck, FileText, Download, Loader2, AlertCircle, CheckCircle2, History, ClipboardCheck } from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import { reconcileFiles, generateExcel, generateCsv, ReconciliationResult, parseExcelReport, parseCsvReport, parseKrcFile, enrichReportWithKrc } from '@/lib/reconciliation';
import { Tabs, TabsContent, TabsList, TabsTrigger } from '@/components/ui/tabs';

interface Stats {
  confirmed: number;
  unconfirmed: number;
  krcChecks: number;
  totalMileage: number;
  totalPlannedMileage: number;
}

export default function App() {
  const [prilFile, setPrilFile] = useState<File | null>(null);
  const [transFile, setTransFile] = useState<File | null>(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [stats, setStats] = useState<Stats | null>(null);
  const [results, setResults] = useState<ReconciliationResult[] | null>(null);
  const [metadata, setMetadata] = useState<{ route: string; month: string; year: string } | null>(null);
  const [tripDuration, setTripDuration] = useState<number>(120);
  const [forwardMileage, setForwardMileage] = useState<number>(0);
  const [returnMileage, setReturnMileage] = useState<number>(0);
  const [krcFile, setKrcFile] = useState<File | null>(null);
  const [reportFile, setReportFile] = useState<File | null>(null);
  const [activeTab, setActiveTab] = useState('reconcile');

  const handleProcess = async () => {
    if (!prilFile || !transFile) {
      setError('Пожалуйста, выберите оба файла для сверки (Приложение и Транзакции).');
      return;
    }

    setLoading(true);
    setError(null);
    setStats(null);
    setResults(null);
    setMetadata(null);

    try {
      const { results: reconResults, stats: reconStats, metadata: reconMeta } = await reconcileFiles(prilFile, transFile, tripDuration, null, forwardMileage, returnMileage);
      setStats({
        ...reconStats,
        totalPlannedMileage: reconResults.reduce((sum, r) => sum + (r.plannedMileage || 0), 0)
      });
      setResults(reconResults);
      setMetadata(reconMeta);
    } catch (err) {
      console.error('Reconciliation error:', err);
      setError(err instanceof Error ? err.message : 'Произошла ошибка при обработке файлов локально');
    } finally {
      setLoading(false);
    }
  };

  const handleKrcStage2 = async () => {
    if (!krcFile) {
      setError('Пожалуйста, выберите файл Отчета КРС.');
      return;
    }

    if (!results || results.length === 0) {
      setError('Сначала выполните первый этап сверки (Сформируйте итоговый отчет).');
      return;
    }

    setLoading(true);
    setError(null);

    try {
      // Use results already in memory from Stage 1
      const reportResults = [...results];

      // 2. Parse KRC data
      const krcData = await parseKrcFile(krcFile);
      if (krcData.length === 0) {
        throw new Error("Не удалось найти данные в файле КРС.");
      }

      // 3. Enrich with KRC
      const enriched = await enrichReportWithKrc(reportResults, krcData);
      
      const confirmed = enriched.filter(r => r.status === 'Подтверждено').length;
      const krcCount = enriched.filter(r => r.krcStatus.includes('время проверки')).length;
      const totalMileage = enriched.reduce((sum, r) => sum + r.mileage, 0);
      const totalPlannedMileage = enriched.reduce((sum, r) => sum + (r.plannedMileage || 0), 0);

      setResults(enriched);
      setStats({
        confirmed,
        unconfirmed: enriched.length - confirmed,
        krcChecks: krcCount,
        totalMileage,
        totalPlannedMileage
      });
      setMetadata({ route: metadata?.route || "из отчета", month: metadata?.month || "", year: metadata?.year || "" });
    } catch (err) {
      console.error('KRC Stage 2 error:', err);
      setError(err instanceof Error ? err.message : 'Ошибка при объединении с КРС');
    } finally {
      setLoading(false);
    }
  };

  const handleDownload = async () => {
    if (!results || !metadata) return;

    try {
      setLoading(true);
      // Stage 1 default is CSV now as requested, Stage 2 (enriched) will still use Excel but can be switched
      const isStage2 = activeTab === 'krc';
      const blob = isStage2 ? await generateExcel(results) : await generateCsv(results);
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      const extension = isStage2 ? 'xlsx' : 'csv';
      const filename = metadata.route !== "неизвестно" 
        ? `сверка по маршруту ${metadata.route} за ${metadata.month} ${metadata.year}.${extension}`
        : `Отчет.${extension}`;
      a.download = filename;
      document.body.appendChild(a);
      a.click();
      window.URL.revokeObjectURL(url);
      document.body.removeChild(a);
    } catch (err) {
      setError('Ошибка при генерации файла');
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="flex min-h-screen bg-[#f8fbfe] text-[#1d3644] font-sans overflow-hidden relative">
      {/* Background Bus Illustration (Simplified SVG) */}
      <div className="absolute bottom-0 right-0 w-[60%] h-[60%] opacity-[0.05] pointer-events-none z-0">
        <svg viewBox="0 0 1000 400" fill="currentColor" className="w-full h-full">
          <path d="M100,300 L900,300 L900,150 C900,120 880,100 850,100 L150,100 C120,100 100,120 100,150 L100,300 Z" />
          <rect x="150" y="130" width="150" height="80" rx="5" fill="#fff" />
          <rect x="320" y="130" width="150" height="80" rx="5" fill="#fff" />
          <rect x="490" y="130" width="150" height="80" rx="5" fill="#fff" />
          <rect x="660" y="130" width="150" height="80" rx="5" fill="#fff" />
          <circle cx="250" cy="300" r="40" stroke="currentColor" strokeWidth="10" fill="none" />
          <circle cx="750" cy="300" r="40" stroke="currentColor" strokeWidth="10" fill="none" />
        </svg>
      </div>

      {/* Sidebar */}
      <aside className="w-[300px] bg-[#e1f0f7] p-8 flex flex-col gap-10 shrink-0 z-10 border-r border-[#c9dde9]">
        {/* Brand Logo Section */}
        <div className="flex flex-col gap-6 py-2">
          <div className="flex items-center gap-3">
            {/* Transit Line Icon */}
            <div className="flex flex-col items-center justify-between h-10 w-2 shrink-0">
              <div className="w-2.5 h-2.5 rounded-full border-2 border-red-600 bg-white" />
              <div className="w-0.5 flex-1 bg-red-600 mx-auto" />
              <div className="w-2.5 h-2.5 rounded-full border-2 border-red-600 bg-white" />
            </div>
            <div className="flex flex-col leading-tight">
              <span className="text-[13px] font-bold text-[#1d3644]">Транспорт</span>
              <span className="text-[15px] font-black text-[#1d3644] -mt-0.5">Верхневолжья</span>
            </div>
          </div>

          <div className="flex items-center gap-3 ml-1">
            {/* Volga Waves Icon */}
            <div className="flex flex-col gap-[3px] shrink-0">
              {[1, 2, 3].map((i) => (
                <svg key={i} width="24" height="6" viewBox="0 0 24 6" fill="none" xmlns="http://www.w3.org/2000/svg">
                  <path d="M2 4C3.5 4 4.5 2 6 2C7.5 2 8.5 4 10 4C11.5 4 12.5 2 14 2C15.5 2 16.5 4 18 4C19.5 4 20.5 2 22 2" stroke="#1d3644" strokeWidth="2.5" strokeLinecap="round" />
                </svg>
              ))}
            </div>
            <span className="text-3xl font-light tracking-[0.15em] text-[#1d3644] ml-1">ВОЛГА</span>
          </div>
        </div>

        <div className="space-y-4">
          <div className="bg-[#b8d9e9] p-5 rounded-xl border-b-[8px] border-[#0076b3] shadow-sm">
            <div className="text-[10px] uppercase tracking-wider text-[#1d3644] font-bold mb-1 opacity-70">Всего рейсов</div>
            <div className="text-2xl font-bold">
              {stats ? stats.confirmed + stats.unconfirmed : '1200'}
            </div>
          </div>

          <div className="bg-[#b8d9e9] p-5 rounded-xl border-b-[8px] border-[#4bb34b] shadow-sm">
            <div className="text-[10px] uppercase tracking-wider text-[#1d3644] font-bold mb-1 opacity-70">Подтверждено</div>
            <div className="text-2xl font-bold">
              {stats ? stats.confirmed : '1150'}
            </div>
          </div>

          <div className="bg-[#b8d9e9] p-5 rounded-xl border-b-[8px] border-[#e23e3e] shadow-sm">
            <div className="text-[10px] uppercase tracking-wider text-[#1d3644] font-bold mb-1 opacity-70">Не подтверждено</div>
            <div className="text-2xl font-bold">
              {stats ? stats.unconfirmed : '50'}
            </div>
          </div>

          <div className="bg-[#b8d9e9] p-5 rounded-xl border-b-[8px] border-[#0070BA] shadow-sm">
            <div className="text-[10px] uppercase tracking-wider text-[#1d3644] font-bold mb-1 opacity-70">Проверок КРС</div>
            <div className="text-2xl font-bold">
              {stats ? stats.krcChecks : '0'}
            </div>
          </div>

          <div className="bg-[#b8d9e9] p-5 rounded-xl border-b-[8px] border-[#00aeef] shadow-sm">
            <div className="text-[10px] uppercase tracking-wider text-[#1d3644] font-bold mb-1 opacity-70">Подтвержденный пробег</div>
            <div className="text-2xl font-bold">
              {stats ? `${stats.totalMileage.toFixed(2)} км` : '0.00 км'}
            </div>
          </div>

          <div className="bg-[#b8d9e9] p-5 rounded-xl border-b-[8px] border-[#648191] shadow-sm">
            <div className="text-[10px] uppercase tracking-wider text-[#1d3644] font-bold mb-1 opacity-70">Плановый пробег</div>
            <div className="text-2xl font-bold">
              {stats ? `${stats.totalPlannedMileage.toFixed(2)} км` : '0.00 км'}
            </div>
          </div>
        </div>

        <div className="mt-auto">
          <div className="text-[10px] uppercase tracking-wider text-[#648191] font-bold">Последний отчет</div>
          <div className="text-sm mt-1 text-[#1d3644]">
            {results && metadata ? (metadata.route !== "неизвестно" ? `рейс ${metadata.route} ${metadata.month}` : 'Отчет готов') : 'Нет данных'}
          </div>
        </div>
      </aside>

      {/* Main Content */}
      <main className="flex-1 p-16 flex flex-col items-start max-w-5xl relative z-10">
        <header className="mb-10 w-full text-left">
          <h1 className="text-4xl font-bold mb-2 text-[#1d3644]">Сверка данных</h1>
          <p className="text-[#648191]">Выберите тип сверки и загрузите файлы для формирования отчета.</p>
        </header>

        <Tabs defaultValue="reconcile" className="w-full" onValueChange={setActiveTab}>
          <TabsList className="bg-[#e1f0f7] p-1 h-14 rounded-2xl mb-8 border border-[#c9dde9]">
            <TabsTrigger 
              value="reconcile" 
              className="rounded-xl px-8 h-full font-bold text-[#1d3644] data-[state=active]:bg-white data-[state=active]:shadow-sm transition-all"
            >
              Сверка рейсов
            </TabsTrigger>
            <TabsTrigger 
              value="krc" 
              className="rounded-xl px-8 h-full font-bold text-[#1d3644] data-[state=active]:bg-white data-[state=active]:shadow-sm transition-all"
            >
              Отчет KRC
            </TabsTrigger>
          </TabsList>

          <TabsContent value="reconcile" className="m-0 space-y-8">
            <div className="grid grid-cols-1 md:grid-cols-2 gap-8 w-full">
              {/* Pril File Dropzone */}
              <div 
                className={`relative h-[220px] rounded-[32px] flex flex-col items-center justify-center p-8 text-center transition-all cursor-pointer ${prilFile ? 'bg-white border-2 border-[#4bb34b] shadow-lg' : 'bg-[#e1f0f7] border-none hover:bg-[#d5eaf5]'}`}
                onClick={() => document.getElementById('pril')?.click()}
              >
                <input id="pril" type="file" accept=".csv" className="hidden" onChange={(e) => setPrilFile(e.target.files?.[0] || null)} />
                <div className="relative mb-4">
                  <div className="w-12 h-16 bg-[#1d3644] rounded-lg -rotate-3 flex items-center justify-center text-white">
                    <FileText className="w-8 h-8 opacity-40" />
                  </div>
                  {prilFile && <div className="absolute -top-2 -right-2 bg-[#4bb34b] rounded-full p-1"><CheckCircle2 className="w-4 h-4 text-white" /></div>}
                </div>
                <strong className="block text-sm text-[#1d3644]">Подтверждаемые рейсы (pril.csv)</strong>
                <p className="text-xs text-[#648191] mt-2">{prilFile ? prilFile.name : 'Перетащите файл или нажмите для выбора'}</p>
              </div>

              {/* Transactions File Dropzone */}
              <div 
                className={`relative h-[220px] rounded-[32px] flex flex-col items-center justify-center p-8 text-center transition-all cursor-pointer ${transFile ? 'bg-white border-2 border-[#0076b3] shadow-lg' : 'bg-[#e1f0f7] border-none hover:bg-[#d5eaf5]'}`}
                onClick={() => document.getElementById('trans')?.click()}
              >
                <input id="trans" type="file" accept=".csv" className="hidden" onChange={(e) => setTransFile(e.target.files?.[0] || null)} />
                <div className="relative mb-4">
                  <div className="w-12 h-16 bg-[#1d3644] rounded-lg rotate-3 flex items-center justify-center text-white">
                    <FileCheck className="w-8 h-8 opacity-40" />
                  </div>
                  {transFile && <div className="absolute -top-2 -right-2 bg-[#0076b3] rounded-full p-1"><CheckCircle2 className="w-4 h-4 text-white" /></div>}
                </div>
                <strong className="block text-sm text-[#1d3644]">Реестр транзакций (transactions.csv)</strong>
                <p className="text-xs text-[#648191] mt-2">{transFile ? transFile.name : 'Перетащите файл или нажмите для выбора'}</p>
              </div>
            </div>

            <div className="flex gap-8 w-full">
              <div className="flex-1 bg-white border-[3px] border-[#00aeef] rounded-[24px] p-8 shadow-sm overflow-hidden flex flex-col items-start gap-3">
                <Label htmlFor="duration" className="text-xs font-bold uppercase tracking-wider text-[#1d3644] opacity-80">Длительность рейса (минуты)</Label>
                <div className="flex items-center gap-4 w-full">
                  <Input id="duration" type="number" value={tripDuration} onChange={(e) => setTripDuration(parseInt(e.target.value) || 0)} className="font-bold text-xl h-14 border-none bg-[#eaf4f9] px-6 rounded-2xl flex-1 focus-visible:ring-0" />
                  <span className="text-[#1d3644] text-lg font-bold">мин</span>
                </div>
                <p className="text-[10px] text-[#648191] mt-1 italic">Окно поиска транзакций после времени начала рейса.</p>
              </div>

              <div className="flex-1 bg-white border-[3px] border-[#4bb34b] rounded-[24px] p-8 shadow-sm overflow-hidden flex flex-col items-start gap-3">
                <Label htmlFor="forward" className="text-xs font-bold uppercase tracking-wider text-[#1d3644] opacity-80">Прямое направление (км)</Label>
                <div className="flex items-center gap-4 w-full">
                  <Input id="forward" type="number" step="0.01" value={forwardMileage} onChange={(e) => setForwardMileage(parseFloat(e.target.value) || 0)} className="font-bold text-xl h-14 border-none bg-[#eaf4f9] px-6 rounded-2xl flex-1 focus-visible:ring-0" />
                  <span className="text-[#1d3644] text-lg font-bold">км</span>
                </div>
                <p className="text-[10px] text-[#648191] mt-1 italic">Плановый пробег для прямого направления.</p>
              </div>

              <div className="flex-1 bg-white border-[3px] border-[#e23e3e] rounded-[24px] p-8 shadow-sm overflow-hidden flex flex-col items-start gap-3">
                <Label htmlFor="return" className="text-xs font-bold uppercase tracking-wider text-[#1d3644] opacity-80">Обратное направление (км)</Label>
                <div className="flex items-center gap-4 w-full">
                  <Input id="return" type="number" step="0.01" value={returnMileage} onChange={(e) => setReturnMileage(parseFloat(e.target.value) || 0)} className="font-bold text-xl h-14 border-none bg-[#eaf4f9] px-6 rounded-2xl flex-1 focus-visible:ring-0" />
                  <span className="text-[#1d3644] text-lg font-bold">км</span>
                </div>
                <p className="text-[10px] text-[#648191] mt-1 italic">Плановый пробег для обратного направления.</p>
              </div>
            </div>
          </TabsContent>

          <TabsContent value="krc" className="m-0 space-y-8">
            <div className="flex justify-between items-center mb-2">
              <h3 className="text-xl font-bold text-[#1d3644]">Этап 2: Интеграция данных КРС</h3>
              {results && (
                <div className="text-xs text-[#648191] bg-white px-3 py-1 rounded-full border border-[#e1f0f7]">
                  Найдено записей: <span className="text-[#1d3644] font-bold">{results.length}</span>
                </div>
              )}
            </div>
            
            <div className="grid grid-cols-1 gap-8 w-full max-w-xl">
              {/* KRC File Dropzone */}
              <div 
                className={`relative h-[240px] rounded-[32px] flex flex-col items-center justify-center p-8 text-center transition-all cursor-pointer ${krcFile ? 'bg-white border-2 border-[#0070BA] shadow-lg' : 'bg-[#e1f0f7] border-none hover:bg-[#d5eaf5]'}`}
                onClick={() => document.getElementById('krc')?.click()}
              >
                <input id="krc" type="file" accept=".csv,.xlsx" className="hidden" onChange={(e) => setKrcFile(e.target.files?.[0] || null)} />
                <div className="relative mb-4">
                  <div className="w-12 h-16 bg-[#1d3644] rounded-lg flex items-center justify-center text-white">
                    <ClipboardCheck className="w-8 h-8 opacity-40" />
                  </div>
                  {krcFile && <div className="absolute -top-2 -right-2 bg-[#0070BA] rounded-full p-1"><CheckCircle2 className="w-4 h-4 text-white" /></div>}
                </div>
                <strong className="block text-sm text-[#1d3644]">Загрузите Отчет KRC</strong>
                <p className="text-xs text-[#648191] mt-2 line-clamp-1">{krcFile ? krcFile.name : 'Данные о проверках КРС применятся к результатам 1 этапа'}</p>
              </div>
            </div>

            <div className="flex flex-col gap-4">
               <Button
                onClick={handleKrcStage2}
                disabled={loading || !krcFile || !results}
                className="w-full md:w-auto px-10 h-16 rounded-[20px] font-bold text-lg bg-[#0070BA] hover:bg-[#005f9e] text-white shadow-md transition-all disabled:opacity-50"
              >
                {loading ? <><Loader2 className="mr-2 h-5 w-5 animate-spin" /> Обработка...</> : 'Объединить с данными КРС'}
              </Button>
              {!results && !loading && (
                <p className="text-xs text-[#e23e3e] font-bold">⚠️ Сначала сформируйте отчет на вкладке "Сверка рейсов"</p>
              )}
            </div>
          </TabsContent>
        </Tabs>

        <div className="flex items-center gap-10 mt-12 w-full">
          {activeTab === 'reconcile' && (
            <Button
              onClick={handleProcess}
              disabled={loading || !prilFile || !transFile}
              className="px-12 h-20 rounded-[24px] font-extrabold text-xl bg-[#0070BA] hover:bg-[#005f9e] text-white shadow-lg transition-all disabled:opacity-50"
            >
              {loading ? (
                <>
                  <Loader2 className="mr-3 h-6 w-6 animate-spin" /> Обработка...
                </>
              ) : (
                'Сформировать промежуточный отчет'
              )}
            </Button>
          )}

          {results && (
            <Button
              onClick={handleDownload}
              className="px-8 h-14 rounded-[20px] font-bold text-lg bg-[#4bb34b] hover:bg-[#3d913d] text-white shadow-md transition-all flex items-center gap-2"
            >
              <Download className="w-5 h-5" /> Скачать отчет
            </Button>
          )}

          <button
            onClick={() => {
              if (activeTab === 'reconcile') {
                setPrilFile(null);
                setTransFile(null);
                setStats(null);
                setResults(null);
                setMetadata(null);
              } else {
                setKrcFile(null);
                setReportFile(null);
                setResults(null);
                setStats(null);
              }
              setError(null);
            }}
            className="text-[#648191] font-bold text-lg hover:text-[#1d3644] transition-colors"
          >
            Очистить
          </button>
        </div>

        <AnimatePresence>
          {error && (
            <motion.div
              initial={{ opacity: 0, y: 10 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0 }}
              className="mt-6 w-full"
            >
              <Alert variant="destructive" className="rounded-xl border-[#e23e3e]/20 bg-red-50 text-[#e23e3e]">
                <AlertCircle className="h-4 w-4" />
                <AlertTitle className="font-bold">Ошибка</AlertTitle>
                <AlertDescription className="text-xs">{error}</AlertDescription>
              </Alert>
            </motion.div>
          )}
        </AnimatePresence>

        <footer className="w-full mt-auto pt-10 flex justify-between text-xs text-[#648191]">
          <span>Статус: {loading ? 'Обработка данных...' : results ? 'Отчет готов' : 'Ожидание файлов'}</span>
          <span>Лицензия: Корпоративная</span>
        </footer>
      </main>
    </div>
  );
}
