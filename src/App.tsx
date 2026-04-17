import { useState } from 'react';
import { Button } from '@/components/ui/button';
import { Card, CardContent, CardDescription, CardFooter, CardHeader, CardTitle } from '@/components/ui/card';
import { Input } from '@/components/ui/input';
import { Label } from '@/components/ui/label';
import { Alert, AlertDescription, AlertTitle } from '@/components/ui/alert';
import { Progress } from '@/components/ui/progress';
import { FileCheck, FileText, Download, Loader2, AlertCircle, CheckCircle2, History, ClipboardCheck } from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import { reconcileFiles, generateExcel, ReconciliationResult } from '@/lib/reconciliation';
import { Tabs, TabsContent, TabsList, TabsTrigger } from '@/components/ui/tabs';

interface Stats {
  confirmed: number;
  unconfirmed: number;
  krcChecks: number;
  totalMileage: number;
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
  const [krcFile, setKrcFile] = useState<File | null>(null);
  const [activeTab, setActiveTab] = useState('reconcile');

  const handleProcess = async () => {
    if (!prilFile || !transFile) {
      setError('Пожалуйста, выберите оба файла для сверки.');
      return;
    }

    setLoading(true);
    setError(null);
    setStats(null);
    setResults(null);
    setMetadata(null);

    try {
      // Process files locally in the browser
      const { results: reconResults, stats: reconStats, metadata: reconMeta } = await reconcileFiles(prilFile, transFile, tripDuration, krcFile);
      
      setStats(reconStats);
      setResults(reconResults);
      setMetadata(reconMeta);
    } catch (err) {
      console.error('Reconciliation error:', err);
      setError(err instanceof Error ? err.message : 'Произошла ошибка при обработке файлов локально');
    } finally {
      setLoading(false);
    }
  };

  const handleDownload = async () => {
    if (!results || !metadata) return;

    try {
      setLoading(true);
      const blob = await generateExcel(results);
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      const filename = metadata.route !== "неизвестно" 
        ? `сверка по маршруту ${metadata.route} за ${metadata.month} ${metadata.year}.xlsx`
        : `Отчет.xlsx`;
      a.download = filename;
      document.body.appendChild(a);
      a.click();
      window.URL.revokeObjectURL(url);
      document.body.removeChild(a);
    } catch (err) {
      setError('Ошибка при генерации Excel-файла');
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

            <div className="w-[450px] bg-white border-[3px] border-[#00aeef] rounded-[24px] p-8 shadow-sm overflow-hidden flex flex-col items-start gap-3">
              <Label htmlFor="duration" className="text-xs font-bold uppercase tracking-wider text-[#1d3644] opacity-80">Длительность рейса (минуты)</Label>
              <div className="flex items-center gap-4 w-full">
                <Input id="duration" type="number" value={tripDuration} onChange={(e) => setTripDuration(parseInt(e.target.value) || 0)} className="font-bold text-xl h-14 border-none bg-[#eaf4f9] px-6 rounded-2xl flex-1 focus-visible:ring-0" />
                <span className="text-[#1d3644] text-lg font-bold">мин</span>
              </div>
              <p className="text-[10px] text-[#648191] mt-1 italic">Окно поиска транзакций после времени начала рейса.</p>
            </div>
          </TabsContent>

          <TabsContent value="krc" className="m-0 space-y-8">
            <div className="w-full">
              {/* KRC File Dropzone */}
              <div 
                className={`relative h-[280px] rounded-[40px] flex flex-col items-center justify-center p-12 text-center transition-all cursor-pointer ${krcFile ? 'bg-white border-2 border-[#0070BA] shadow-lg' : 'bg-[#e1f0f7] border-none hover:bg-[#d5eaf5]'}`}
                onClick={() => document.getElementById('krc')?.click()}
              >
                <input id="krc" type="file" accept=".csv,.xlsx" className="hidden" onChange={(e) => setKrcFile(e.target.files?.[0] || null)} />
                <div className="relative mb-6">
                  <div className="w-16 h-20 bg-[#1d3644] rounded-xl flex items-center justify-center text-white">
                    <ClipboardCheck className="w-10 h-10 opacity-40" />
                  </div>
                  {krcFile && <div className="absolute -top-2 -right-2 bg-[#0070BA] rounded-full p-2 shadow-sm"><CheckCircle2 className="w-5 h-5 text-white" /></div>}
                </div>
                <h3 className="text-xl font-bold text-[#1d3644] mb-2">Отчет KRC</h3>
                <p className="text-[#648191] max-w-sm mx-auto">
                  {krcFile ? krcFile.name : 'Загрузите файл отчета KRC для интеграции данных в общую сверку.'}
                </p>
                <div className="mt-6 px-4 py-2 bg-white/50 rounded-full text-[10px] font-bold text-[#0070BA] uppercase tracking-wider">
                  Поддерживаемые форматы: .CSV, .XLSX
                </div>
              </div>

              {krcFile && (
                 <motion.div 
                  initial={{ opacity: 0, y: 10 }} 
                  animate={{ opacity: 1, y: 0 }}
                  className="mt-8 p-6 bg-white border-2 border-[#e1f0f7] rounded-[32px] flex items-center gap-4"
                >
                  <div className="p-3 bg-[#e1f0f7] rounded-2xl">
                    <History className="w-6 h-6 text-[#0070BA]" />
                  </div>
                  <div>
                    <h4 className="font-bold text-[#1d3644]">Файл готов к обработке</h4>
                    <p className="text-sm text-[#648191]">Нажмите на кнопку ниже, чтобы начать анализ данных KRC.</p>
                  </div>
                </motion.div>
              )}
            </div>
          </TabsContent>
        </Tabs>

        <div className="flex items-center gap-10 mt-12 w-full">
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
              'Сформировать итоговый отчет'
            )}
          </Button>

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
