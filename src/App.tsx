import { useEffect, useMemo, useState } from 'react';
import ExcelJS from 'exceljs';
import {
  ArrowDownToLine,
  Copy,
  FileSpreadsheet,
  Mail,
  Search,
  Settings,
  ShieldAlert,
  TrendingUp,
  Upload,
} from 'lucide-react';
import {
  createColumnHelper,
  flexRender,
  getCoreRowModel,
  getFilteredRowModel,
  getSortedRowModel,
  useReactTable,
} from '@tanstack/react-table';
import {
  Bar,
  BarChart,
  CartesianGrid,
  Cell,
  Legend,
  Pie,
  PieChart,
  ResponsiveContainer,
  Tooltip,
  XAxis,
  YAxis,
} from 'recharts';

type ClientRow = {
  ID_CLIENTE: string;
  NOME_CLIENTE: string;
  RESPONSAVEL: string;
};

type NfRow = {
  ID_CLIENTE: string;
  NF: string;
  DATA_EMISSAO: string;
  CFOP: string;
};

type PortalRow = {
  OBJETO: string;
  SERVICO: string;
  PESO: string;
  QTD: string;
  POSTAGEM: string;
  VALOR: string;
  DECLARADO: string;
  A_COBRAR: string;
  DESTINATARIO: string;
  CEP: string;
  SITUACAO: string;
  DATA_SITUACAO: string;
  NF: string;
  DEPARTAMENTO: string;
  ADICIONAIS: string;
  CONTEUDO: string;
  CONTRATO_ECT: string;
  DESTINO: string;
  CODIGO_PP: string;
  LISTA_POSTAGEM: string;
  CIDADE: string;
  UF: string;
  PRAZO_ESTIMADO: string;
  PRAZO_REAL: string;
  OBS: string;
  ALTURA: string;
  LARGURA: string;
  COMPRIMENTO: string;
  CODIGO_ECT: string;
  CARTAO_POSTAGEM: string;
  PLP: string;
  RFID: string;
};

type ProcessedRow = {
  NF: string;
  DATA_DA_POSTAGEM: string;
  DESTINATARIO: string;
  SITUACAO: string;
  CFOP: string;
  AMOSTRA_OU_VENDA: string;
  ID: string;
  RESPONSAVEL: string;
  DIAS_PARADOS: number;
  PRIORIDADE: string;
  CLIENTE: string;
  UF: string;
  DATA_POSTAGEM_RAW: string;
};

type SpecialConfig = {
  cliente: string;
  nomeResponsavel: string;
  mensagem: string;
};

type FilterState = {
  responsavel: string;
  cliente: string;
  situacao: string;
  estado: string;
  tipoOperacao: string;
  faixaDias: string;
  dataPostagem: string;
};

const initialFilters: FilterState = {
  responsavel: 'all',
  cliente: 'all',
  situacao: 'all',
  estado: 'all',
  tipoOperacao: 'all',
  faixaDias: 'all',
  dataPostagem: '',
};

const columnHelper = createColumnHelper<ProcessedRow>();

const normalizeValue = (value: unknown) => {
  if (value === null || value === undefined) return '';
  return String(value).trim();
};

const worksheetToJson = <T extends Record<string, string>>(worksheet: ExcelJS.Worksheet) => {
  const headerRow = worksheet.getRow(1);
  const headers = Array.from({ length: headerRow.cellCount }, (_, index) =>
    normalizeValue(headerRow.getCell(index + 1).text),
  );

  const rows: T[] = [];
  worksheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return;

    const entry: Record<string, string> = {};
    let hasValues = false;

    headers.forEach((header, index) => {
      if (!header) return;

      const value = normalizeValue(row.getCell(index + 1).text);
      entry[header] = value;
      if (value) hasValues = true;
    });

    if (hasValues) {
      rows.push(entry as T);
    }
  });

  return rows;
};

const exportColumns = [
  { header: 'NF', key: 'NF' },
  { header: 'DATA DA POSTAGEM', key: 'DATA_DA_POSTAGEM' },
  { header: 'DESTINATÁRIO', key: 'DESTINATARIO' },
  { header: 'SITUAÇÃO', key: 'SITUACAO' },
  { header: 'CFOP', key: 'CFOP' },
  { header: 'AMOSTRA OU VENDA', key: 'AMOSTRA_OU_VENDA' },
  { header: 'ID', key: 'ID' },
  { header: 'RESPONSÁVEL', key: 'RESPONSAVEL' },
  { header: 'DIAS_PARADOS', key: 'DIAS_PARADOS' },
  { header: 'PRIORIDADE', key: 'PRIORIDADE' },
] as const;

const escapeCsvValue = (value: string | number) => {
  const normalized = String(value).replaceAll('"', '""');
  return /[",;\n]/.test(normalized) ? `"${normalized}"` : normalized;
};

const parseDate = (value: string) => {
  if (!value) return null;
  const [day, month, year] = value.split(/[/-]/).map((part) => part.trim());
  if (!day || !month || !year) return null;
  const parsed = new Date(`${year}-${month.padStart(2, '0')}-${day.padStart(2, '0')}`);
  return Number.isNaN(parsed.getTime()) ? null : parsed;
};

const priorityByDays = (days: number) => {
  if (days > 15) return '🔴 CRÍTICO';
  if (days >= 8) return '🟡 ATENÇÃO';
  return '🟢 NORMAL';
};

const specialAlertSituations = [
  'Objeto aguardando retirada no endereço indicado',
  'Entrega não realizada',
  'Destinatário ausente',
  'Endereço incorreto',
  'Objeto devolvido',
  'Tentativa de entrega não efetuada',
];

const App = () => {
  const [rows, setRows] = useState<ProcessedRow[]>([]);
  const [summary, setSummary] = useState<Record<string, number>>({});
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState('');
  const [filters, setFilters] = useState<FilterState>(initialFilters);
  const [specialConfigs, setSpecialConfigs] = useState<SpecialConfig[]>([]);
  const [emailBody, setEmailBody] = useState('');
  const [emailHtml, setEmailHtml] = useState('');

  useEffect(() => {
    setSpecialConfigs(loadSpecialConfigs());
  }, []);

  const handleFileUpload = async (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;

    setLoading(true);
    setError('');
    try {
      const data = await file.arrayBuffer();
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(data);
      const requiredSheets = ['Clientes', 'NFS FATURADAS', 'Portal Postal'];
      const missingSheets = requiredSheets.filter((sheet) => !workbook.getWorksheet(sheet));
      if (missingSheets.length) {
        throw new Error(`Aba(s) ausente(s): ${missingSheets.join(', ')}`);
      }

      const clientsSheet = worksheetToJson<ClientRow>(workbook.getWorksheet('Clientes')!);
      const nfsSheet = worksheetToJson<NfRow>(workbook.getWorksheet('NFS FATURADAS')!);
      const portalSheet = worksheetToJson<PortalRow>(workbook.getWorksheet('Portal Postal')!);

      const clientsMap = new Map<string, ClientRow>();
      clientsSheet.forEach((client) => {
        const id = normalizeValue(client.ID_CLIENTE);
        if (id) clientsMap.set(id, client);
      });

      const nfsMap = new Map<string, NfRow[]>();
      nfsSheet.forEach((nf) => {
        const id = normalizeValue(nf.ID_CLIENTE);
        if (!id) return;
        const existing = nfsMap.get(id) ?? [];
        existing.push(nf);
        nfsMap.set(id, existing);
      });

      const processed: ProcessedRow[] = [];
      portalSheet.forEach((portal) => {
        const nf = normalizeValue(portal.NF);
        if (!nf) return;

        const matchingNf = nfsSheet.find((item) => normalizeValue(item.NF) === nf);
        if (!matchingNf) return;

        const client = clientsMap.get(normalizeValue(matchingNf.ID_CLIENTE));
        const postingDate = parseDate(normalizeValue(portal.POSTAGEM));
        const today = new Date();
        const daysParados = postingDate
          ? Math.max(0, Math.floor((today.getTime() - postingDate.getTime()) / (1000 * 60 * 60 * 24)))
          : 0;

        const cfop = normalizeValue(matchingNf.CFOP);
        const tipoOperacao = cfop === '5911' || cfop === '6911' ? 'AMOSTRA' : 'VENDA';

        processed.push({
          NF: nf,
          DATA_DA_POSTAGEM: normalizeValue(portal.POSTAGEM),
          DESTINATARIO: normalizeValue(portal.DESTINATARIO),
          SITUACAO: normalizeValue(portal.SITUACAO),
          CFOP: cfop,
          AMOSTRA_OU_VENDA: tipoOperacao,
          ID: normalizeValue(matchingNf.ID_CLIENTE),
          RESPONSAVEL: normalizeValue(client?.RESPONSAVEL ?? ''),
          DIAS_PARADOS: daysParados,
          PRIORIDADE: priorityByDays(daysParados),
          CLIENTE: normalizeValue(client?.NOME_CLIENTE ?? ''),
          UF: normalizeValue(portal.UF),
          DATA_POSTAGEM_RAW: normalizeValue(portal.POSTAGEM),
        });
      });

      const sortedRows = processed.sort((a, b) => {
        if (b.DIAS_PARADOS !== a.DIAS_PARADOS) return b.DIAS_PARADOS - a.DIAS_PARADOS;
        return a.DATA_POSTAGEM_RAW.localeCompare(b.DATA_POSTAGEM_RAW);
      });

      setRows(sortedRows);
      setSummary({
        total: sortedRows.length,
        nfs: new Set(sortedRows.map((row) => row.NF)).size,
        vendas: sortedRows.filter((row) => row.AMOSTRA_OU_VENDA === 'VENDA').length,
        amostras: sortedRows.filter((row) => row.AMOSTRA_OU_VENDA === 'AMOSTRA').length,
        clientes: new Set(sortedRows.map((row) => row.CLIENTE)).size,
        responsaveis: new Set(sortedRows.map((row) => row.RESPONSAVEL)).size,
        criticos: sortedRows.filter((row) => row.PRIORIDADE === '🔴 CRÍTICO').length,
      });
      setSpecialConfigs(loadSpecialConfigs());
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Erro ao processar o arquivo.');
      setRows([]);
      setSummary({});
    } finally {
      setLoading(false);
    }
  };

  const filteredRows = useMemo(() => {
    return rows.filter((row) => {
      const matchesResponsavel = filters.responsavel === 'all' || row.RESPONSAVEL === filters.responsavel;
      const matchesCliente = filters.cliente === 'all' || row.CLIENTE === filters.cliente;
      const matchesSituacao = filters.situacao === 'all' || row.SITUACAO === filters.situacao;
      const matchesEstado = filters.estado === 'all' || row.UF === filters.estado;
      const matchesTipo = filters.tipoOperacao === 'all' || row.AMOSTRA_OU_VENDA === filters.tipoOperacao;
      const matchesData = !filters.dataPostagem || row.DATA_POSTAGEM_RAW.includes(filters.dataPostagem);
      let matchesDias = true;
      if (filters.faixaDias !== 'all') {
        const [min, max] = filters.faixaDias.split('-').map(Number);
        matchesDias = row.DIAS_PARADOS >= min && row.DIAS_PARADOS <= max;
      }
      return matchesResponsavel && matchesCliente && matchesSituacao && matchesEstado && matchesTipo && matchesDias && matchesData;
    });
  }, [rows, filters]);

  const chartData = useMemo(() => {
    const byResponsavel = rows.reduce<Record<string, number>>((acc, row) => {
      acc[row.RESPONSAVEL || 'Sem responsável'] = (acc[row.RESPONSAVEL || 'Sem responsável'] || 0) + 1;
      return acc;
    }, {});
    const bySituacao = rows.reduce<Record<string, number>>((acc, row) => {
      acc[row.SITUACAO || 'Sem situação'] = (acc[row.SITUACAO || 'Sem situação'] || 0) + 1;
      return acc;
    }, {});
    const byEstado = rows.reduce<Record<string, number>>((acc, row) => {
      acc[row.UF || 'Sem UF'] = (acc[row.UF || 'Sem UF'] || 0) + 1;
      return acc;
    }, {});
    const byTipo = rows.reduce<Record<string, number>>((acc, row) => {
      acc[row.AMOSTRA_OU_VENDA] = (acc[row.AMOSTRA_OU_VENDA] || 0) + 1;
      return acc;
    }, {});
    const topClients = Object.entries(rows.reduce<Record<string, number>>((acc, row) => {
      acc[row.CLIENTE || 'Sem cliente'] = (acc[row.CLIENTE || 'Sem cliente'] || 0) + 1;
      return acc;
    }, {})).sort((a, b) => b[1] - a[1]).slice(0, 10);
    const byDate = rows.reduce<Record<string, number>>((acc, row) => {
      const key = row.DATA_DA_POSTAGEM || 'Sem data';
      acc[key] = (acc[key] || 0) + 1;
      return acc;
    }, {});

    return {
      responsavel: Object.entries(byResponsavel).map(([name, value]) => ({ name, value })),
      situacao: Object.entries(bySituacao).map(([name, value]) => ({ name, value })),
      estado: Object.entries(byEstado).map(([name, value]) => ({ name, value })),
      tipo: Object.entries(byTipo).map(([name, value]) => ({ name, value })),
      topClients: topClients.map(([name, value]) => ({ name, value })),
      byDate: Object.entries(byDate).map(([name, value]) => ({ name, value })),
    };
  }, [rows]);

  const columns = useMemo(() => [
    columnHelper.accessor('NF', { id: 'NF', header: 'NF' }),
    columnHelper.accessor('DATA_DA_POSTAGEM', { id: 'DATA_DA_POSTAGEM', header: 'DATA DA POSTAGEM' }),
    columnHelper.accessor('DESTINATARIO', { id: 'DESTINATARIO', header: 'DESTINATÁRIO' }),
    columnHelper.accessor('SITUACAO', { id: 'SITUACAO', header: 'SITUAÇÃO' }),
    columnHelper.accessor('CFOP', { id: 'CFOP', header: 'CFOP' }),
    columnHelper.accessor('AMOSTRA_OU_VENDA', { id: 'AMOSTRA_OU_VENDA', header: 'AMOSTRA OU VENDA' }),
    columnHelper.accessor('ID', { id: 'ID', header: 'ID' }),
    columnHelper.accessor('RESPONSAVEL', { id: 'RESPONSAVEL', header: 'RESPONSÁVEL' }),
    columnHelper.accessor('DIAS_PARADOS', { id: 'DIAS_PARADOS', header: 'DIAS PARADOS' }),
    columnHelper.accessor('PRIORIDADE', { id: 'PRIORIDADE', header: 'PRIORIDADE' }),
  ], []);

  const table = useReactTable({
    data: filteredRows,
    columns,
    getCoreRowModel: getCoreRowModel(),
    getFilteredRowModel: getFilteredRowModel(),
    getSortedRowModel: getSortedRowModel(),
    initialState: {
      sorting: [{ id: 'DIAS_PARADOS', desc: true }],
    },
  });

  const generateEmail = () => {
    const specialMessages = specialConfigs
      .filter((config) => config.cliente && filteredRows.some((row) => row.ID === config.cliente))
      .map((config) => `@${config.nomeResponsavel}, ${config.mensagem}`);

    const tableHtml = `
      <table style="border-collapse: collapse; width: 100%; font-family: Arial, sans-serif; font-size: 12px; margin-top: 12px;">
        <thead>
          <tr style="background: #0f3b6d; color: white;">
            <th style="border: 1px solid #cbd5e1; padding: 8px; text-align: left;">NF</th>
            <th style="border: 1px solid #cbd5e1; padding: 8px; text-align: left;">DATA DA POSTAGEM</th>
            <th style="border: 1px solid #cbd5e1; padding: 8px; text-align: left;">DESTINATÁRIO</th>
            <th style="border: 1px solid #cbd5e1; padding: 8px; text-align: left;">SITUAÇÃO</th>
            <th style="border: 1px solid #cbd5e1; padding: 8px; text-align: left;">CFOP</th>
            <th style="border: 1px solid #cbd5e1; padding: 8px; text-align: left;">AMOSTRA OU VENDA</th>
            <th style="border: 1px solid #cbd5e1; padding: 8px; text-align: left;">ID</th>
            <th style="border: 1px solid #cbd5e1; padding: 8px; text-align: left;">RESPONSÁVEL</th>
          </tr>
        </thead>
        <tbody>
          ${filteredRows.map((row) => `
            <tr style="background: #f8fafc;">
              <td style="border: 1px solid #cbd5e1; padding: 8px;">${row.NF}</td>
              <td style="border: 1px solid #cbd5e1; padding: 8px;">${row.DATA_DA_POSTAGEM}</td>
              <td style="border: 1px solid #cbd5e1; padding: 8px;">${row.DESTINATARIO}</td>
              <td style="border: 1px solid #cbd5e1; padding: 8px;">${row.SITUACAO}</td>
              <td style="border: 1px solid #cbd5e1; padding: 8px;">${row.CFOP}</td>
              <td style="border: 1px solid #cbd5e1; padding: 8px;">${row.AMOSTRA_OU_VENDA}</td>
              <td style="border: 1px solid #cbd5e1; padding: 8px;">${row.ID}</td>
              <td style="border: 1px solid #cbd5e1; padding: 8px;">${row.RESPONSAVEL}</td>
            </tr>
          `).join('')}
        </tbody>
      </table>
    `;

    const body = `Segue abaixo a lista dos materiais parados nos Correios, poderiam por favor avisar os clientes urgente para evitar devolução?

A orientação da gerência é focar nos casos mais antigos, realizar a comunicação via telefone e formalizar por e-mail após contato. Caso o cliente não atenda, formalizar diretamente o e-mail.

Observação Importante:
Nossa modalidade de envio de Sedex é para entregar direto na planta do cliente, porém, algumas vezes, principalmente por motivos de geolocalização.

${specialMessages.join('\n\n')}

${tableHtml}

Obrigado!

Atenciosamente,`;

    const htmlBody = `
      <div style="font-family: Arial, sans-serif; color: #0f172a; line-height: 1.5;">
        <p>Segue abaixo a lista dos materiais parados nos Correios, poderiam por favor avisar os clientes urgente para evitar devolução?</p>
        <p>A orientação da gerência é focar nos casos mais antigos, realizar a comunicação via telefone e formalizar por e-mail após contato. Caso o cliente não atenda, formalizar diretamente o e-mail.</p>
        <p><strong>Observação Importante:</strong><br/>Nossa modalidade de envio de Sedex é para entregar direto na planta do cliente, porém, algumas vezes, principalmente por motivos de geolocalização.</p>
        ${specialMessages.length ? `<p>${specialMessages.join('<br/><br/>')}</p>` : ''}
        ${tableHtml}
        <p>Obrigado!</p>
        <p>Atenciosamente,</p>
      </div>
    `;

    setEmailBody(body);
    setEmailHtml(htmlBody);
  };

  const copyEmail = async () => {
    if (!emailHtml) {
      generateEmail();
    }
    try {
      await navigator.clipboard.writeText(emailHtml);
      window.alert('E-mail copiado para a área de transferência em HTML.');
    } catch {
      window.alert('Não foi possível copiar o e-mail.');
    }
  };

  const exportTable = async (type: 'xlsx' | 'csv' | 'pdf') => {
    if (type === 'pdf') {
      const printWindow = window.open('', '_blank', 'width=900,height=700');
      if (!printWindow) return;
      const html = `
        <html>
          <head><title>Relatório de Materiais Parados</title></head>
          <body style="font-family: Arial; padding: 24px;">
            <h2>Relatório de Materiais Parados</h2>
            <table style="border-collapse: collapse; width: 100%; font-size: 12px;">
              <thead>
                <tr style="background: #0f3b6d; color: white;">
                  <th style="border: 1px solid #cbd5e1; padding: 6px; text-align: left;">NF</th>
                  <th style="border: 1px solid #cbd5e1; padding: 6px; text-align: left;">DATA DA POSTAGEM</th>
                  <th style="border: 1px solid #cbd5e1; padding: 6px; text-align: left;">DESTINATÁRIO</th>
                  <th style="border: 1px solid #cbd5e1; padding: 6px; text-align: left;">SITUAÇÃO</th>
                  <th style="border: 1px solid #cbd5e1; padding: 6px; text-align: left;">CFOP</th>
                  <th style="border: 1px solid #cbd5e1; padding: 6px; text-align: left;">AMOSTRA OU VENDA</th>
                  <th style="border: 1px solid #cbd5e1; padding: 6px; text-align: left;">RESPONSÁVEL</th>
                </tr>
              </thead>
              <tbody>
                ${filteredRows.map((row) => `
                  <tr>
                    <td style="border: 1px solid #cbd5e1; padding: 6px;">${row.NF}</td>
                    <td style="border: 1px solid #cbd5e1; padding: 6px;">${row.DATA_DA_POSTAGEM}</td>
                    <td style="border: 1px solid #cbd5e1; padding: 6px;">${row.DESTINATARIO}</td>
                    <td style="border: 1px solid #cbd5e1; padding: 6px;">${row.SITUACAO}</td>
                    <td style="border: 1px solid #cbd5e1; padding: 6px;">${row.CFOP}</td>
                    <td style="border: 1px solid #cbd5e1; padding: 6px;">${row.AMOSTRA_OU_VENDA}</td>
                    <td style="border: 1px solid #cbd5e1; padding: 6px;">${row.RESPONSAVEL}</td>
                  </tr>
                `).join('')}
              </tbody>
            </table>
          </body>
        </html>
      `;
      printWindow.document.write(html);
      printWindow.document.close();
      printWindow.print();
      return;
    }

    const exportData = filteredRows.map((row) => ({
      NF: row.NF,
      DATA_DA_POSTAGEM: row.DATA_DA_POSTAGEM,
      DESTINATARIO: row.DESTINATARIO,
      SITUACAO: row.SITUACAO,
      CFOP: row.CFOP,
      AMOSTRA_OU_VENDA: row.AMOSTRA_OU_VENDA,
      ID: row.ID,
      RESPONSAVEL: row.RESPONSAVEL,
      DIAS_PARADOS: row.DIAS_PARADOS,
      PRIORIDADE: row.PRIORIDADE,
    }));

    const blob = type === 'csv'
      ? new Blob([
        `\uFEFF${[
          exportColumns.map((column) => escapeCsvValue(column.header)).join(';'),
          ...exportData.map((row) => exportColumns
            .map((column) => escapeCsvValue(row[column.key]))
            .join(';')),
        ].join('\n')}`,
      ], { type: 'text/csv;charset=utf-8;' })
      : new Blob([await (async () => {
        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('Relatório');
        worksheet.columns = exportColumns.map((column) => ({
          header: column.header,
          key: column.key,
          width: Math.max(column.header.length + 2, 18),
        }));
        worksheet.addRows(exportData);
        return workbook.xlsx.writeBuffer();
      })()]);
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = `relatorio.${type === 'csv' ? 'csv' : 'xlsx'}`;
    link.click();
    URL.revokeObjectURL(url);
  };

  const copyReport = () => {
    navigator.clipboard.writeText(filteredRows.map((row) => `${row.NF}\t${row.DATA_DA_POSTAGEM}\t${row.DESTINATARIO}`).join('\n'));
  };

  const saveSpecialConfigs = () => {
    const configText = JSON.stringify(specialConfigs, null, 2);
    localStorage.setItem('special-configs', configText);
    window.alert('Configurações salvas localmente.');
  };

  const loadSpecialConfigs = (): SpecialConfig[] => {
    const stored = localStorage.getItem('special-configs');
    return stored ? JSON.parse(stored) : [];
  };

  const updateConfig = (index: number, field: keyof SpecialConfig, value: string) => {
    setSpecialConfigs((current) => {
      const next = [...current];
      next[index] = { ...next[index], [field]: value };
      return next;
    });
  };

  const addConfig = () => {
    setSpecialConfigs((current) => [...current, { cliente: '', nomeResponsavel: '', mensagem: '' }]);
  };

  return (
    <div className="min-h-screen bg-slate-950 text-slate-100">
      <div className="mx-auto max-w-7xl p-6 space-y-6">
        <header className="rounded-3xl border border-slate-800 bg-slate-900/80 p-8 shadow-2xl shadow-slate-950/40">
          <div className="flex flex-col gap-4 lg:flex-row lg:items-end lg:justify-between">
            <div>
              <p className="text-sm uppercase tracking-[0.35em] text-sky-400">Correios • Monitor de materiais parados</p>
              <h1 className="mt-2 text-3xl font-semibold text-white">Automação profissional para análise e comunicação</h1>
              <p className="mt-3 max-w-3xl text-sm text-slate-300">
                Carregue o Excel, o sistema processa as abas, gera o dashboard, organiza a tabela e monta o e-mail corporativo em segundos.
              </p>
            </div>
            <label className="flex cursor-pointer items-center gap-3 rounded-2xl border border-sky-500/40 bg-sky-500/10 px-5 py-3 text-sm font-medium text-sky-300 transition hover:bg-sky-500/20">
              <Upload size={18} />
              <span>Upload de Excel</span>
              <input type="file" accept=".xlsx,.xls" onChange={handleFileUpload} className="hidden" />
            </label>
          </div>
        </header>

        {error ? (
          <div className="rounded-2xl border border-rose-500/40 bg-rose-500/10 p-4 text-sm text-rose-200">{error}</div>
        ) : null}

        {loading ? <div className="rounded-2xl border border-slate-800 bg-slate-900/70 p-6 text-slate-200">Processando arquivo...</div> : null}

        {!rows.length && !loading && !error ? (
          <div className="rounded-3xl border border-dashed border-slate-700 bg-slate-900/70 p-12 text-center text-slate-300">
            <FileSpreadsheet className="mx-auto mb-3" size={40} />
            <p className="text-lg font-semibold text-white">Pronto para começar</p>
            <p className="mt-2 text-sm">Faça upload do Excel para visualizar o dashboard e gerar o e-mail.</p>
          </div>
        ) : null}

        {rows.length ? (
          <>
            <section className="grid gap-4 md:grid-cols-2 xl:grid-cols-4">
              {[
                { label: 'Total de Materiais Parados', value: summary.total ?? 0, icon: ShieldAlert },
                { label: 'Quantidade de NFs', value: summary.nfs ?? 0, icon: FileSpreadsheet },
                { label: 'Quantidade de Vendas', value: summary.vendas ?? 0, icon: TrendingUp },
                { label: 'Quantidade de Amostras', value: summary.amostras ?? 0, icon: Mail },
              ].map((card) => {
                const Icon = card.icon;
                return (
                  <div key={card.label} className="rounded-3xl border border-slate-800 bg-slate-900/70 p-5 shadow-lg shadow-slate-950/20">
                    <div className="flex items-center justify-between">
                      <p className="text-sm text-slate-400">{card.label}</p>
                      <div className="rounded-2xl bg-sky-500/10 p-2 text-sky-400">
                        <Icon size={18} />
                      </div>
                    </div>
                    <p className="mt-6 text-3xl font-semibold text-white">{card.value}</p>
                  </div>
                );
              })}
            </section>

            <section className="grid gap-4 xl:grid-cols-2">
              <div className="rounded-3xl border border-slate-800 bg-slate-900/70 p-5">
                <div className="mb-4 flex items-center justify-between">
                  <h2 className="text-lg font-semibold text-white">Filtros</h2>
                  <Search size={18} className="text-slate-400" />
                </div>
                <div className="grid gap-3 md:grid-cols-2">
                  {[
                    { key: 'responsavel', label: 'Responsável', options: ['all', ...Array.from(new Set(rows.map((row) => row.RESPONSAVEL))).filter(Boolean)] },
                    { key: 'cliente', label: 'Cliente', options: ['all', ...Array.from(new Set(rows.map((row) => row.CLIENTE))).filter(Boolean)] },
                    { key: 'situacao', label: 'Situação', options: ['all', ...Array.from(new Set(rows.map((row) => row.SITUACAO))).filter(Boolean)] },
                    { key: 'estado', label: 'Estado', options: ['all', ...Array.from(new Set(rows.map((row) => row.UF))).filter(Boolean)] },
                    { key: 'tipoOperacao', label: 'Tipo de Operação', options: ['all', 'AMOSTRA', 'VENDA'] },
                    { key: 'faixaDias', label: 'Faixa de Dias Parados', options: ['all', '0-7', '8-15', '16-999'] },
                  ].map((filter) => (
                    <label key={filter.key} className="flex flex-col gap-2 text-sm text-slate-300">
                      <span>{filter.label}</span>
                      <select
                        value={filters[filter.key as keyof FilterState] as string}
                        onChange={(event) => setFilters((current) => ({ ...current, [filter.key]: event.target.value }))}
                        className="rounded-xl border border-slate-700 bg-slate-950/80 px-3 py-2"
                      >
                        {filter.options.map((option) => (
                          <option key={option} value={option}>{option === 'all' ? 'Todos' : option}</option>
                        ))}
                      </select>
                    </label>
                  ))}
                  <label className="flex flex-col gap-2 text-sm text-slate-300">
                    <span>Data de Postagem</span>
                    <input
                      type="text"
                      value={filters.dataPostagem}
                      onChange={(event) => setFilters((current) => ({ ...current, dataPostagem: event.target.value }))}
                      className="rounded-xl border border-slate-700 bg-slate-950/80 px-3 py-2"
                      placeholder="Ex.: 2024-05-20"
                    />
                  </label>
                </div>
              </div>

              <div className="rounded-3xl border border-slate-800 bg-slate-900/70 p-5">
                <div className="mb-4 flex items-center justify-between">
                  <h2 className="text-lg font-semibold text-white">Configuração administrativa</h2>
                  <Settings size={18} className="text-slate-400" />
                </div>
                <div className="space-y-3">
                  {specialConfigs.map((config, index) => (
                    <div key={`${config.cliente}-${index}`} className="rounded-2xl border border-slate-800 bg-slate-950/70 p-3">
                      <div className="grid gap-3 md:grid-cols-3">
                        <input
                          value={config.cliente}
                          onChange={(event) => updateConfig(index, 'cliente', event.target.value)}
                          placeholder="Cliente"
                          className="rounded-xl border border-slate-700 bg-slate-900 px-3 py-2"
                        />
                        <input
                          value={config.nomeResponsavel}
                          onChange={(event) => updateConfig(index, 'nomeResponsavel', event.target.value)}
                          placeholder="Nome Responsável"
                          className="rounded-xl border border-slate-700 bg-slate-900 px-3 py-2"
                        />
                        <input
                          value={config.mensagem}
                          onChange={(event) => updateConfig(index, 'mensagem', event.target.value)}
                          placeholder="Mensagem"
                          className="rounded-xl border border-slate-700 bg-slate-900 px-3 py-2"
                        />
                      </div>
                    </div>
                  ))}
                  <div className="flex gap-3">
                    <button onClick={addConfig} className="rounded-xl border border-slate-700 bg-slate-800 px-3 py-2 text-sm">Adicionar cliente especial</button>
                    <button onClick={saveSpecialConfigs} className="rounded-xl bg-sky-600 px-3 py-2 text-sm font-semibold text-white">Salvar configurações</button>
                  </div>
                </div>
              </div>
            </section>

            <section className="grid gap-4 xl:grid-cols-2">
              <div className="rounded-3xl border border-slate-800 bg-slate-900/70 p-5">
                <h2 className="mb-4 text-lg font-semibold text-white">Dashboard executivo</h2>
                <div className="h-72">
                  <ResponsiveContainer width="100%" height="100%">
                    <BarChart data={chartData.responsavel}>
                      <CartesianGrid strokeDasharray="3 3" stroke="#334155" />
                      <XAxis dataKey="name" tick={{ fill: '#cbd5e1' }} />
                      <YAxis tick={{ fill: '#cbd5e1' }} />
                      <Tooltip />
                      <Legend />
                      <Bar dataKey="value" fill="#38bdf8" />
                    </BarChart>
                  </ResponsiveContainer>
                </div>
              </div>
              <div className="rounded-3xl border border-slate-800 bg-slate-900/70 p-5">
                <h2 className="mb-4 text-lg font-semibold text-white">Materiais por situação</h2>
                <div className="h-72">
                  <ResponsiveContainer width="100%" height="100%">
                    <PieChart>
                      <Pie data={chartData.situacao} dataKey="value" nameKey="name" outerRadius={90} fill="#0ea5e9">
                        {chartData.situacao.map((entry, index) => (
                          <Cell key={`${entry.name}-${index}`} fill={['#38bdf8', '#818cf8', '#f59e0b', '#f472b6'][index % 4]} />
                        ))}
                      </Pie>
                      <Tooltip />
                      <Legend />
                    </PieChart>
                  </ResponsiveContainer>
                </div>
              </div>
            </section>

            <section className="grid gap-4 xl:grid-cols-2">
              <div className="rounded-3xl border border-slate-800 bg-slate-900/70 p-5">
                <h2 className="mb-4 text-lg font-semibold text-white">Top 10 clientes</h2>
                <div className="h-72">
                  <ResponsiveContainer width="100%" height="100%">
                    <BarChart data={chartData.topClients}>
                      <CartesianGrid strokeDasharray="3 3" stroke="#334155" />
                      <XAxis dataKey="name" tick={{ fill: '#cbd5e1' }} />
                      <YAxis tick={{ fill: '#cbd5e1' }} />
                      <Tooltip />
                      <Bar dataKey="value" fill="#34d399" />
                    </BarChart>
                  </ResponsiveContainer>
                </div>
              </div>
              <div className="rounded-3xl border border-slate-800 bg-slate-900/70 p-5">
                <h2 className="mb-4 text-lg font-semibold text-white">Evolução por data</h2>
                <div className="h-72">
                  <ResponsiveContainer width="100%" height="100%">
                    <BarChart data={chartData.byDate}>
                      <CartesianGrid strokeDasharray="3 3" stroke="#334155" />
                      <XAxis dataKey="name" tick={{ fill: '#cbd5e1' }} />
                      <YAxis tick={{ fill: '#cbd5e1' }} />
                      <Tooltip />
                      <Bar dataKey="value" fill="#fb923c" />
                    </BarChart>
                  </ResponsiveContainer>
                </div>
              </div>
            </section>

            <section className="rounded-3xl border border-slate-800 bg-slate-900/70 p-5">
              <div className="mb-4 flex flex-col gap-3 lg:flex-row lg:items-center lg:justify-between">
                <div>
                  <h2 className="text-lg font-semibold text-white">Tabela operacional</h2>
                  <p className="text-sm text-slate-400">Visual semelhante a um modelo corporativo de e-mail com ordenação crítica.</p>
                </div>
                <div className="flex flex-wrap gap-3">
                  <button onClick={generateEmail} className="rounded-xl bg-slate-700 px-3 py-2 text-sm font-semibold text-white">Gerar E-mail</button>
                  <button onClick={copyEmail} className="flex items-center gap-2 rounded-xl bg-sky-600 px-3 py-2 text-sm font-semibold text-white"><Copy size={16} /> Copiar E-mail</button>
                  <button onClick={() => exportTable('xlsx')} className="flex items-center gap-2 rounded-xl bg-emerald-600 px-3 py-2 text-sm font-semibold text-white"><ArrowDownToLine size={16} /> Exportar Excel</button>
                  <button onClick={() => exportTable('csv')} className="rounded-xl bg-amber-600 px-3 py-2 text-sm font-semibold text-white">Exportar CSV</button>
                  <button onClick={() => exportTable('pdf')} className="rounded-xl bg-rose-600 px-3 py-2 text-sm font-semibold text-white">Exportar PDF</button>
                  <button onClick={copyReport} className="rounded-xl bg-violet-600 px-3 py-2 text-sm font-semibold text-white">Copiar Relatório</button>
                </div>
              </div>

              <div className="overflow-x-auto rounded-2xl border border-slate-800">
                <table className="min-w-full border-collapse text-sm">
                  <thead className="bg-sky-900 text-slate-100">
                    {table.getHeaderGroups().map((headerGroup) => (
                      <tr key={headerGroup.id}>
                        {headerGroup.headers.map((header) => (
                          <th key={header.id} className="border border-slate-700 px-3 py-2 text-left">
                            {flexRender(header.column.columnDef.header, header.getContext())}
                          </th>
                        ))}
                      </tr>
                    ))}
                  </thead>
                  <tbody>
                    {table.getRowModel().rows.map((row, index) => {
                      const isAlert = specialAlertSituations.includes(row.original.SITUACAO);
                      return (
                        <tr key={row.id} className={`border-b border-slate-800 ${index % 2 === 0 ? 'bg-slate-950/60' : 'bg-slate-900/60'} ${isAlert ? 'text-rose-300' : 'text-slate-200'}`}>
                          {row.getVisibleCells().map((cell) => (
                            <td key={cell.id} className="border border-slate-800 px-3 py-2">
                              {flexRender(cell.column.columnDef.cell, cell.getContext())}
                            </td>
                          ))}
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            </section>

            {emailBody ? (
              <section className="rounded-3xl border border-slate-800 bg-slate-900/70 p-5">
                <div className="mb-3 flex items-center justify-between">
                  <h2 className="text-lg font-semibold text-white">E-mail corporativo</h2>
                  <button onClick={copyEmail} className="rounded-xl bg-sky-600 px-3 py-2 text-sm font-semibold text-white">Copiar para Outlook</button>
                </div>
                <div className="rounded-2xl border border-slate-800 bg-white p-4 text-slate-900">
                  <div className="prose prose-sm max-w-none" dangerouslySetInnerHTML={{ __html: emailHtml }} />
                </div>
              </section>
            ) : null}
          </>
        ) : null}
      </div>
    </div>
  );
};

export default App;
