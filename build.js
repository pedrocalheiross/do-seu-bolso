// Lê dados/*.xlsx (uma planilha por farmácia, abas "resumo" e "ofertas") e gera precos.json.
const fs = require('fs');
const assert = require('assert');
const XLSX = require('xlsx');

const POR_DOSE = ['comprimidos', 'cápsulas', 'doses'];

const mediana = (xs) => {
  const s = [...xs].sort((a, b) => a - b), m = s.length >> 1;
  return s.length % 2 ? s[m] : (s[m - 1] + s[m]) / 2;
};
const arred = (x) => Math.round(x * 10000) / 10000;

assert.strictEqual(mediana([3, 1, 2]), 2);
assert.strictEqual(mediana([1, 2, 3, 4]), 2.5);

const produtos = {};
const fontes = [];

for (const arq of fs.readdirSync('dados').filter((f) => f.endsWith('.xlsx'))) {
  const farmacia = arq.replace('.xlsx', '');
  const wb = XLSX.readFile(`dados/${arq}`);
  const ler = (aba) => XLSX.utils.sheet_to_json(wb.Sheets[aba]);

  if (wb.Sheets.notas) {
    const linha = XLSX.utils.sheet_to_json(wb.Sheets.notas, { header: 1 }).flat().find((t) => /coletado/i.test(t));
    fontes.push({ farmacia, nota: linha || '' });
  }

  // "resumo" traz o elenco inteiro, inclusive itens sem oferta.
  for (const r of ler('resumo')) {
    // imagem_url vem dos sites das farmácias e vai direto para innerHTML no index.html.
    assert(!r.imagem_url || /^https:\/\/[^"'<>\s]+$/.test(r.imagem_url), `imagem_url suspeita em ${arq}: ${r.imagem_url}`);
    produtos[r.produto] ??= { produto: r.produto, indicacao: r.indicacao, imagem_url: r.imagem_url || null, ofertas: [] };
    if (r.imagem_url) produtos[r.produto].imagem_url = r.imagem_url;
  }

  for (const o of ler('ofertas')) {
    if (o.no_programa !== 'sim' || o.disponivel_recife === 'não') continue;
    produtos[o.produto_programa] ??= { produto: o.produto_programa, indicacao: o.indicacao, imagem_url: null, ofertas: [] };
    produtos[o.produto_programa].ofertas.push({
      unidade: o.unidade, preco_caixa: o.preco_1_unidade, preco_unidade: o.preco_por_unidade,
    });
  }
}

const saida = Object.values(produtos).map((p) => {
  const unidades = p.ofertas.map((o) => o.unidade);
  const principal = unidades.sort((a, b) => unidades.filter((u) => u === b).length - unidades.filter((u) => u === a).length)[0];
  // Anticoncepcional em cartela se compra por cartela/mês, não por comprimido/dia.
  const modo = POR_DOSE.includes(principal) && !/CARTELA/.test(p.produto) ? 'dose' : 'caixa';
  // ponytail: ofertas sem unidade (ex.: "drágeas") ficam fora do modo dose; parsear "apresentacao" se fizer falta.
  const precos = p.ofertas
    .filter((o) => (modo === 'dose' ? POR_DOSE.includes(o.unidade) && o.preco_unidade : o.preco_caixa))
    .map((o) => arred(modo === 'dose' ? o.preco_unidade : o.preco_caixa))
    .sort((a, b) => a - b);
  return {
    produto: p.produto, indicacao: p.indicacao, imagem_url: p.imagem_url || null, modo,
    unidade: modo === 'dose' ? principal : 'caixa/frasco',
    ...(precos.length
      ? { min: precos[0], mediana: arred(mediana(precos)), max: precos.at(-1) }
      : { sem_preco: true }),
  };
}).sort((a, b) => a.indicacao.localeCompare(b.indicacao) || a.produto.localeCompare(b.produto));

const metformina = saida.find((p) => p.produto === 'CLORIDRATO DE METFORMINA 850MG');
assert.strictEqual(metformina.modo, 'dose');
assert.strictEqual(metformina.min, 0.172);

fs.writeFileSync('precos.json', JSON.stringify({ gerado_em: new Date().toISOString().slice(0, 10), fontes, produtos: saida }, null, 1));
console.log(`${saida.length} produtos (${saida.filter((p) => p.sem_preco).length} sem preço) -> precos.json`);
