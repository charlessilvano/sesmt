# Gerador de SST para Condomínios

Versão 4, completa e compatível com GitHub Pages, do aplicativo PWA para preparar documentos de Saúde e Segurança do Trabalho de condomínios. O fluxo reúne PGR, relatório psicossocial da NR-01, PCMSO e LTCAT, com geração dos arquivos em Word (`.docx`) diretamente no navegador.

## Publicação no GitHub Pages

1. Extraia este ZIP.
2. Crie um repositório vazio no GitHub.
3. Envie **todo o conteúdo desta pasta**, inclusive a pasta `.github`.
4. Use a branch `main`.
5. No repositório, abra **Settings → Pages**.
6. Em **Build and deployment**, escolha **GitHub Actions**.
7. Abra a aba **Actions** e acompanhe a rotina **Publicar no GitHub Pages**.

Após a conclusão, o endereço publicado aparecerá na própria execução e em **Settings → Pages**. Novos envios para a branch `main` serão publicados automaticamente.

## Executar no computador

É necessário ter Node.js 22 e pnpm 11.

```bash
pnpm install
pnpm dev
```

Para testar o mesmo pacote utilizado pelo GitHub Pages:

```bash
pnpm build
pnpm preview
```

## Recursos incluídos

- seleção dos condomínios do documento-base;
- configuração do quadro de funcionários por cargo e GHE;
- checklist de atividades, ambientes e exposições;
- inventário de riscos com matriz qualitativa 6 × 4;
- EPIs, treinamentos, inspeções e plano de ação;
- importação local do CSV de respostas da NR-01;
- vinculação automática por nome ou CNPJ do condomínio;
- indicadores por domínio, questões prioritárias e plano de ação psicossocial;
- proteção de resultados detalhados quando houver menos de três respostas;
- inclusão do relatório NR-01 como anexo no mesmo arquivo Word do PGR;
- elaboração prioritária do PCMSO a partir dos cargos e riscos configurados no PGR;
- protocolos médicos por função, com a descrição da atividade idêntica à do PGR e todos os riscos detalhados por natureza, fonte, agravos e situação;
- matriz de 423 vínculos função–exame extraídos da planilha de referência, cobrindo 25 funções e 32 exames ou avaliações;
- detalhamento dos eventos admissional, seis meses após admissão, periódico, mudança de risco, retorno ao trabalho e demissional;
- distinção explícita entre exame-base, condicional ao risco, indicação individual, cenário não caracterizado, cuidado pós-exposição e parte do exame clínico;
- campos do médico responsável, CRM, RQE, clínica, vigência, vacinação e relatório analítico;
- elaboração do LTCAT para todas as funções e grupos de trabalhadores, inclusive quando não há agente nocivo caracterizado;
- registro das condições ambientais, dos riscos por função e da conclusão previdenciária de cada função;
- avaliações qualitativas ou quantitativas separadas para os agentes previdenciários potencialmente aplicáveis;
- controle de método, resultado, unidade, critério, EPC, EPI, eficácia e conclusão por exposição;
- documentos Word com corpo em Arial 12, espaçamento 1,5, capa com a identidade BIOSAFE e rodapé institucional;
- central de documentos com geração separada de PGR, PCMSO e LTCAT em Word;
- salvamento automático no navegador;
- instalação como PWA e funcionamento básico offline;
- funcionamento estático, sem servidor e sem envio dos dados a terceiros.

## Arquivos gerados

- **PGR:** inventário de riscos, plano de ação e, quando selecionado, relatório NR-01 anexado;
- **PCMSO:** minuta médica em formato paisagem, com atividade, riscos e matriz detalhada de exames por função, vinculados ao PGR;
- **LTCAT:** minuta previdenciária que contempla todas as funções, descreve suas condições ambientais e riscos e registra as avaliações técnicas pendentes.

## Nota técnica

Os arquivos gerados são minutas estruturadas. Cargos, ambientes, atividades, exposições, controles e enquadramentos legais devem ser confirmados em campo. O PGR exige validação do responsável técnico; o PCMSO exige revisão e assinatura do médico responsável; e o LTCAT exige inspeção e assinatura de médico do trabalho ou engenheiro de segurança legalmente habilitado. O questionário fornecido é tratado como instrumento personalizado de consulta e triagem coletiva, não como aplicação integral do HSE Indicator Tool.
