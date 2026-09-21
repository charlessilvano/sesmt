import {
  AlignmentType,
  BorderStyle,
  Document,
  Footer,
  Header,
  HeadingLevel,
  PageBreak,
  PageNumber,
  Packer,
  Paragraph,
  ShadingType,
  Table,
  TableCell,
  TableLayoutType,
  TableRow,
  TextRun,
  VerticalAlign,
  WidthType,
} from "docx";

import { getSelectedRoles } from "./pgr-selectors";
import type { PgrState } from "./pgr-types";
import {
  getLtcatExposureRows,
  getPcmsoProtocolRows,
  ltcatRowIsComplete,
} from "./occupational";

const COLORS = {
  navy: "163B55",
  navySoft: "E9F0F4",
  ink: "17212B",
  muted: "5A6872",
  line: "D9E0E5",
  pale: "F5F7F8",
  white: "FFFFFF",
  green: "DDEFE4",
  yellow: "FFF0BF",
  red: "F8D6D6",
};

const borders = {
  top: { style: BorderStyle.SINGLE, size: 1, color: COLORS.line },
  bottom: { style: BorderStyle.SINGLE, size: 1, color: COLORS.line },
  left: { style: BorderStyle.SINGLE, size: 1, color: COLORS.line },
  right: { style: BorderStyle.SINGLE, size: 1, color: COLORS.line },
  insideHorizontal: { style: BorderStyle.SINGLE, size: 1, color: COLORS.line },
  insideVertical: { style: BorderStyle.SINGLE, size: 1, color: COLORS.line },
};

const cellMargins = { top: 105, bottom: 105, left: 125, right: 125 };

function text(value: unknown, fallback = "Não informado") {
  const normalized = String(value ?? "").trim();
  return normalized || fallback;
}

function date(value: string) {
  if (!value) return "Não informada";
  const [year, month, day] = value.split("-");
  return year && month && day ? `${day}/${month}/${year}` : value;
}

function paragraph(value: string, options: { italic?: boolean; center?: boolean; boldLead?: string } = {}) {
  const runs: TextRun[] = [];
  if (options.boldLead && value.startsWith(options.boldLead)) {
    runs.push(new TextRun({ text: options.boldLead, bold: true }));
    runs.push(new TextRun({ text: value.slice(options.boldLead.length), italics: options.italic }));
  } else {
    runs.push(new TextRun({ text: value, italics: options.italic }));
  }
  return new Paragraph({
    children: runs,
    alignment: options.center ? AlignmentType.CENTER : AlignmentType.JUSTIFIED,
    spacing: { after: 145, line: 300 },
  });
}

function bullet(value: string) {
  return new Paragraph({
    children: [new TextRun(value)],
    bullet: { level: 0 },
    spacing: { after: 85, line: 285 },
  });
}

function heading(value: string) {
  return new Paragraph({
    text: value,
    heading: HeadingLevel.HEADING_1,
    keepNext: true,
    spacing: { before: 250, after: 135 },
  });
}

function subheading(value: string) {
  return new Paragraph({
    text: value,
    heading: HeadingLevel.HEADING_2,
    keepNext: true,
    spacing: { before: 190, after: 100 },
  });
}

function cell(value: string, options: { header?: boolean; fill?: string; width?: number; bold?: boolean } = {}) {
  return new TableCell({
    children: [
      new Paragraph({
        children: [
          new TextRun({
            text: value,
            bold: options.header || options.bold,
            color: options.header ? COLORS.white : COLORS.ink,
            size: 18,
          }),
        ],
        spacing: { after: 0, line: 235 },
      }),
    ],
    borders,
    margins: cellMargins,
    verticalAlign: VerticalAlign.CENTER,
    width: options.width ? { size: options.width, type: WidthType.PERCENTAGE } : undefined,
    shading: options.header
      ? { fill: COLORS.navy, color: "auto", type: ShadingType.CLEAR }
      : options.fill
        ? { fill: options.fill, color: "auto", type: ShadingType.CLEAR }
        : undefined,
  });
}

function table(headers: string[], rows: string[][], widths?: number[]) {
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    layout: TableLayoutType.FIXED,
    borders,
    rows: [
      new TableRow({
        tableHeader: true,
        cantSplit: true,
        children: headers.map((header, index) => cell(header, { header: true, width: widths?.[index] })),
      }),
      ...rows.map(
        (row, rowIndex) =>
          new TableRow({
            cantSplit: true,
            children: row.map((value, columnIndex) =>
              cell(value, {
                width: widths?.[columnIndex],
                fill: rowIndex % 2 ? COLORS.pale : COLORS.white,
              }),
            ),
          }),
      ),
    ],
  });
}

function gap() {
  return new Paragraph({ spacing: { after: 90 } });
}

function cover(
  acronym: string,
  title: string,
  subtitle: string,
  state: PgrState,
  label: string,
) {
  return [
    new Paragraph({
      children: [new TextRun({ text: "GRUPO DINAMIZZA", bold: true, color: COLORS.navy, size: 22, characterSpacing: 80 })],
      alignment: AlignmentType.CENTER,
      spacing: { before: 880, after: 590 },
    }),
    new Paragraph({ text: acronym, style: "Title", alignment: AlignmentType.CENTER, spacing: { after: 120 } }),
    new Paragraph({
      children: [new TextRun({ text: title, bold: true, size: 27 })],
      alignment: AlignmentType.CENTER,
      spacing: { after: 160 },
    }),
    new Paragraph({
      children: [new TextRun({ text: subtitle, color: COLORS.muted, size: 22 })],
      alignment: AlignmentType.CENTER,
      spacing: { after: 400 },
    }),
    new Paragraph({
      children: [new TextRun({ text: text(state.condoName), bold: true, color: COLORS.navy, size: 30 })],
      alignment: AlignmentType.CENTER,
      spacing: { after: 560 },
    }),
    table(
      ["Identificação", "Informação"],
      [
        ["CNPJ", text(state.cnpj)],
        ["Endereço", [text(state.address, ""), text(state.cityState, "")].filter(Boolean).join(" – ") || "Não informado"],
        ["Empresa prestadora", text(state.contractingCompany, "GRUPO DINAMIZZA")],
        ["Data-base", date(state.issueDate)],
      ],
      [32, 68],
    ),
    new Paragraph({ spacing: { before: 420 } }),
    new Paragraph({
      children: [new TextRun({ text: label, bold: true, color: "9A5A14", size: 19 })],
      alignment: AlignmentType.CENTER,
      spacing: { after: 90 },
    }),
    paragraph(
      "O conteúdo deve ser conferido com a realidade do trabalho e assinado pelo profissional legalmente habilitado antes de sua utilização oficial.",
      { italic: true, center: true },
    ),
    new Paragraph({ children: [new PageBreak()] }),
  ];
}

function headerFor(acronym: string, state: PgrState) {
  return new Header({
    children: [
      new Paragraph({
        children: [
          new TextRun({ text: `${acronym}  |  `, bold: true, color: COLORS.navy, size: 17 }),
          new TextRun({ text: text(state.condoName), color: COLORS.muted, size: 17 }),
        ],
        spacing: { after: 70 },
        border: { bottom: { style: BorderStyle.SINGLE, size: 7, color: COLORS.line } },
      }),
    ],
  });
}

function footerFor(acronym: string) {
  return new Footer({
    children: [
      new Paragraph({
        children: [
          new TextRun({ text: `${acronym}  •  página `, color: COLORS.muted, size: 17 }),
          new TextRun({ children: [PageNumber.CURRENT], color: COLORS.muted, size: 17 }),
          new TextRun({ text: " de ", color: COLORS.muted, size: 17 }),
          new TextRun({ children: [PageNumber.TOTAL_PAGES], color: COLORS.muted, size: 17 }),
        ],
        alignment: AlignmentType.CENTER,
      }),
    ],
  });
}

function baseDocument(title: string, description: string, state: PgrState, acronym: string, children: Array<Paragraph | Table>) {
  return new Document({
    creator: text(state.contractingCompany, "GRUPO DINAMIZZA"),
    title,
    subject: description,
    description,
    styles: {
      default: {
        document: { run: { font: "Arial", size: 21, color: COLORS.ink }, paragraph: { spacing: { line: 300 } } },
      },
      paragraphStyles: [
        {
          id: "Title",
          name: "Title",
          basedOn: "Normal",
          next: "Normal",
          quickFormat: true,
          run: { font: "Arial", size: 46, bold: true, color: "000000" },
          paragraph: { spacing: { after: 180 }, alignment: AlignmentType.CENTER },
        },
        {
          id: "Heading1",
          name: "Heading 1",
          basedOn: "Normal",
          next: "Normal",
          quickFormat: true,
          run: { font: "Arial", size: 25, bold: true, color: "000000" },
          paragraph: { spacing: { before: 250, after: 135 }, keepNext: true },
        },
        {
          id: "Heading2",
          name: "Heading 2",
          basedOn: "Normal",
          next: "Normal",
          quickFormat: true,
          run: { font: "Arial", size: 22, bold: true, color: "000000" },
          paragraph: { spacing: { before: 190, after: 100 }, keepNext: true },
        },
      ],
    },
    sections: [
      {
        properties: {
          titlePage: true,
          page: {
            size: { width: 12240, height: 15840 },
            margin: { top: 920, right: 900, bottom: 920, left: 900, header: 420, footer: 420 },
          },
        },
        headers: { default: headerFor(acronym, state), first: new Header({ children: [] }) },
        footers: { default: footerFor(acronym), first: new Footer({ children: [] }) },
        children,
      },
    ],
  });
}

export function buildPcmsoDocument(state: PgrState) {
  const protocols = getPcmsoProtocolRows(state);
  const totalWorkers = getSelectedRoles(state).reduce((sum, role) => sum + role.quantity, 0);
  const children: Array<Paragraph | Table> = [
    ...cover(
      "PCMSO",
      "Programa de Controle Médico de Saúde Ocupacional",
      "Vigilância da saúde integrada aos riscos do PGR",
      state,
      "MINUTA PARA REVISÃO, VALIDAÇÃO E ASSINATURA DO MÉDICO RESPONSÁVEL",
    ),
    heading("1. Identificação do programa"),
    table(
      ["Campo", "Informação"],
      [
        ["Empregador ou estabelecimento", text(state.condoName)],
        ["CNPJ", text(state.cnpj)],
        ["Endereço", text(state.address)],
        ["Cidade e UF", text(state.cityState)],
        ["CNAE", text(state.pcmso.cnae)],
        ["Grau de risco", text(state.pcmso.riskGrade)],
        ["Trabalhadores abrangidos", String(totalWorkers)],
        ["Vigência", `${date(state.pcmso.startDate)} a ${date(state.pcmso.endDate)}`],
        ["Médico responsável", text(state.pcmso.physicianName)],
        ["CRM", text(state.pcmso.physicianCrm)],
        ["RQE", text(state.pcmso.physicianRqe, "Não informado")],
        ["Clínica ou serviço", text(state.pcmso.clinicName, "Não informado")],
      ],
      [31, 69],
    ),
    heading("2. Objetivo e integração com o PGR"),
    paragraph(
      "Este PCMSO estabelece a vigilância ativa e passiva da saúde dos trabalhadores abrangidos, considerando os riscos ocupacionais identificados e classificados no PGR. O programa busca detectar precocemente possíveis agravos relacionados ao trabalho, acompanhar grupos de trabalhadores e apoiar a revisão das medidas de prevenção.",
    ),
    paragraph(
      "A matriz clínica foi produzida a partir do inventário de riscos selecionado no sistema. Ela funciona como proposta técnica para análise do médico responsável, que deve confirmar riscos, critérios, exames, periodicidades e condutas antes da assinatura.",
    ),
    heading("3. Responsabilidades"),
    subheading("3.1 Empregador"),
    bullet("Garantir a elaboração e a implantação do PCMSO, sem custo para o empregado."),
    bullet("Indicar médico responsável e disponibilizar o PGR, o inventário de riscos e as informações das atividades."),
    bullet("Assegurar a realização dos exames ocupacionais e implementar as recomendações decorrentes da vigilância em saúde."),
    subheading("3.2 Médico responsável pelo PCMSO"),
    bullet("Revisar os riscos do PGR e definir os procedimentos médicos adequados a cada grupo de trabalhadores."),
    bullet("Orientar os médicos examinadores, estabelecer critérios de interpretação e emitir ou supervisionar a emissão dos ASOs."),
    bullet("Produzir o relatório analítico anual e comunicar achados que indiquem necessidade de revisão das medidas de prevenção."),
    subheading("3.3 Trabalhadores e gestão local"),
    bullet("Comparecer aos exames e fornecer informações verdadeiras sobre o trabalho e a saúde."),
    bullet("Comunicar acidentes, sintomas, mudanças de atividade e exposições não previstas."),
    heading("4. Exames médicos ocupacionais"),
    table(
      ["Exame", "Momento mínimo de realização"],
      [
        ["Admissional", "Antes de o empregado assumir suas atividades."],
        ["Periódico", "Conforme a periodicidade definida pelo médico, observados os riscos e os critérios da NR-07."],
        ["Retorno ao trabalho", "Antes de reassumir as funções após ausência de 30 dias ou mais por doença ou acidente, ocupacional ou não."],
        ["Mudança de riscos ocupacionais", "Antes da mudança, quando a nova atividade alterar os riscos ocupacionais."],
        ["Demissional", "Em até 10 dias contados do término do contrato, observadas as hipóteses de dispensa previstas na NR-07."],
      ],
      [29, 71],
    ),
    gap(),
    paragraph(
      "Todo exame ocupacional deve incluir avaliação clínica com anamnese ocupacional e exame físico e mental. Exames complementares somente devem ser solicitados quando previstos na NR-07, em seus anexos, ou quando o médico responsável os considerar tecnicamente necessários com base nos riscos e na condição clínica.",
    ),
    heading("5. Matriz de vigilância da saúde por função"),
    paragraph(
      "A periodicidade e os exames abaixo constituem proposta inicial. O médico responsável deve ajustar cada protocolo depois de analisar a exposição real, as avaliações ambientais, as informações clínicas e os requisitos específicos da NR-07.",
    ),
  ];

  if (!protocols.length) {
    children.push(paragraph("Nenhuma função foi incluída. O quadro de trabalhadores deve ser preenchido antes da emissão."));
  }

  protocols.forEach((row, index) => {
    children.push(
      new Paragraph({ children: [new PageBreak()] }),
      subheading(`5.${index + 1} ${row.roleName} – ${row.ghe}`),
      table(
        ["Elemento", "Diretriz proposta"],
        [
          ["Quantidade", `${row.quantity} trabalhador(es)`],
          ["Riscos do PGR", row.riskFactors],
          ["Possíveis agravos", row.possibleHarms],
          ["Avaliação clínica direcionada", row.clinicalFocus],
          ["Exames complementares", row.complementaryExams],
          ["Periodicidade", row.periodicity],
        ],
        [29, 71],
      ),
      gap(),
    );
  });

  children.push(
    new Paragraph({ children: [new PageBreak()] }),
    heading("6. Diretrizes para exames complementares"),
    bullet("Audiometria: aplicar somente quando a exposição a níveis de pressão sonora elevados atender aos critérios da NR-07 e após confirmação da exposição."),
    bullet("Agentes químicos: identificar o agente, a composição, a via de absorção, a dose, a FDS e o critério específico antes de definir exames."),
    bullet("Agentes biológicos: considerar atividade real, incidentes, imunização e protocolos de pós-exposição; não presumir exames laboratoriais de rotina."),
    bullet("Riscos ergonômicos e psicossociais: usar avaliação clínica, sintomas, dados agregados e encaminhamentos indicados, preservando o sigilo médico."),
    bullet("Atividades críticas: a decisão de aptidão pertence ao médico examinador e não pode ser automatizada por cargo ou por checklist."),
    heading("7. Atestado de Saúde Ocupacional"),
    paragraph("Para cada exame clínico ocupacional, o ASO deve registrar, no mínimo:"),
    bullet("razão social e CNPJ ou CAEPF do empregador;"),
    bullet("nome completo, CPF e função do empregado;"),
    bullet("riscos ocupacionais específicos que exijam controle médico, ou a indicação de sua inexistência;"),
    bullet("exames realizados e respectivas datas;"),
    bullet("conclusão de apto ou inapto para a função;"),
    bullet("identificação do médico responsável pelo PCMSO, quando houver;"),
    bullet("data, número de registro e assinatura do médico examinador."),
    heading("8. Condutas, encaminhamentos e emergência"),
    paragraph(
      "Achados clínicos ou complementares possivelmente relacionados ao trabalho devem levar à avaliação médica, ao encaminhamento assistencial e às providências ocupacionais cabíveis. Quando houver suspeita de doença relacionada ao trabalho, devem ser avaliados o registro da CAT, o afastamento da exposição e a revisão do PGR.",
    ),
    paragraph(`Referência para urgência e emergência: ${text(state.pcmso.emergencyReferral, "a definir pelo empregador e pelo médico responsável")}.`, { boldLead: "Referência para urgência e emergência:" }),
    paragraph(`Imunização: ${text(state.pcmso.vaccinationGuidance)}.`, { boldLead: "Imunização:" }),
    heading("9. Relatório analítico anual"),
    paragraph(
      "O médico responsável deve elaborar relatório analítico anual, considerando os dados agregados do programa e preservando o sigilo médico. O relatório deve subsidiar o planejamento do período seguinte e a revisão das medidas de prevenção.",
    ),
    table(
      ["Conteúdo mínimo de acompanhamento", "Registro esperado"],
      [
        ["Exames clínicos", "Número de avaliações por tipo e por grupo de trabalhadores"],
        ["Exames complementares", "Número e tipos de exames realizados"],
        ["Resultados anormais", "Estatística por tipo de exame, sem identificação nominal"],
        ["Doenças relacionadas ao trabalho", "Incidência e prevalência identificadas no período"],
        ["CAT", "Número, tipo de evento e doenças registradas"],
        ["Comparação anual", "Análise em relação ao relatório anterior e definição das ações do período seguinte"],
        ["Data prevista", date(state.pcmso.analyticalReportDueDate)],
      ],
      [38, 62],
    ),
    heading("10. Prontuários e sigilo"),
    paragraph(
      "Os dados dos exames clínicos e complementares devem ser registrados em prontuário médico individual sob responsabilidade do médico responsável. Os prontuários devem ser mantidos por, no mínimo, 20 anos após o desligamento do empregado, salvo disposição específica em contrário. O acesso deve respeitar o sigilo médico e a legislação de proteção de dados.",
    ),
    heading("11. Revisão do programa"),
    bullet("Revisar o PCMSO quando o PGR identificar novos riscos, mudanças de atividade ou alteração relevante da exposição."),
    bullet("Analisar eventos sentinela, acidentes, CAT, afastamentos e resultados agregados dos exames."),
    bullet("Registrar as mudanças de protocolo e orientar os médicos examinadores."),
    paragraph(`Observações do programa: ${text(state.pcmso.notes, "Sem observações adicionais")}.`, { boldLead: "Observações do programa:" }),
    heading("12. Aprovação médica"),
    paragraph(
      "Declaro que revisei os riscos ocupacionais, os protocolos e as periodicidades deste PCMSO e que as definições médicas refletem as condições de trabalho conhecidas na data da assinatura.",
    ),
    new Paragraph({ spacing: { before: 620 } }),
    table(
      ["Médico responsável", "Registro", "Data e assinatura"],
      [[text(state.pcmso.physicianName), [text(state.pcmso.physicianCrm), text(state.pcmso.physicianRqe, "")].filter(Boolean).join(" · ") || "Não informado", "________________________________"]],
      [42, 25, 33],
    ),
    gap(),
    paragraph(
      "Base normativa consultada: NR-07 vigente; NR-01 e PGR do estabelecimento. Conferir eventuais alterações legais antes da assinatura.",
      { italic: true },
    ),
  );

  return baseDocument(
    `PCMSO – ${text(state.condoName)}`,
    "Programa de Controle Médico de Saúde Ocupacional integrado ao PGR",
    state,
    "PCMSO",
    children,
  );
}

export function buildLtcatDocument(state: PgrState) {
  const rows = getLtcatExposureRows(state);
  const selectedRoles = getSelectedRoles(state);
  const pending = rows.filter((row) => !ltcatRowIsComplete(row));
  const children: Array<Paragraph | Table> = [
    ...cover(
      "LTCAT",
      "Laudo Técnico das Condições Ambientais do Trabalho",
      "Caracterização previdenciária das exposições ocupacionais",
      state,
      "MINUTA TÉCNICA — NÃO UTILIZAR NO PPP SEM AVALIAÇÃO E ASSINATURA PROFISSIONAL",
    ),
    heading("1. Identificação"),
    table(
      ["Campo", "Informação"],
      [
        ["Estabelecimento", text(state.condoName)],
        ["CNPJ", text(state.cnpj)],
        ["Endereço", text(state.address)],
        ["Cidade e UF", text(state.cityState)],
        ["Data da avaliação", date(state.ltcat.assessmentDate)],
        ["Responsável técnico", text(state.ltcat.professionalName)],
        ["Habilitação", text(state.ltcat.professionalTitle)],
        ["Registro profissional", text(state.ltcat.professionalRegistration)],
      ],
      [31, 69],
    ),
    heading("2. Objetivo e limites"),
    paragraph(
      "Este LTCAT tem por objetivo registrar as condições ambientais e a efetiva exposição a agentes físicos, químicos e biológicos para fins previdenciários. O documento deve sustentar as informações correspondentes do Perfil Profissiográfico Previdenciário quando estiver concluído e assinado.",
    ),
    paragraph(
      "O sistema reaproveita agentes identificados no PGR, mas não converte riscos trabalhistas em enquadramento previdenciário automático. A conclusão exige inspeção, descrição da exposição, metodologia adequada e, quando aplicável, avaliação quantitativa.",
    ),
    heading("3. Referências legais e técnicas"),
    bullet("Lei nº 8.213/1991, art. 58, quanto à comprovação da exposição e à emissão do laudo por médico do trabalho ou engenheiro de segurança do trabalho."),
    bullet("Decreto nº 3.048/1999, especialmente os arts. 64 a 68 e o Anexo IV."),
    bullet("Instrução Normativa PRES/INSS nº 128/2022 e suas alterações vigentes na data da emissão."),
    bullet("Normas Regulamentadoras e procedimentos de higiene ocupacional aplicáveis a cada agente."),
    heading("4. Metodologia"),
    paragraph(text(state.ltcat.inspectionMethod)),
    table(
      ["Item", "Registro"],
      [
        ["Equipamentos utilizados", text(state.ltcat.equipment, "Não informado ou não aplicável")],
        ["Calibração e rastreabilidade", text(state.ltcat.calibration, "Não informada ou não aplicável")],
        ["Observações técnicas", text(state.ltcat.technicalNotes, "Sem observações adicionais")],
      ],
      [31, 69],
    ),
    heading("5. Funções e grupos avaliados"),
    selectedRoles.length
      ? table(
          ["GHE", "Função", "Quantidade", "Descrição da atividade"],
          selectedRoles.map((role) => [role.ghe, role.name, String(role.quantity), role.description]),
          [14, 24, 12, 50],
        )
      : paragraph("Nenhuma função foi incluída no quadro."),
    heading("6. Avaliação dos agentes"),
    paragraph(
      "As linhas abaixo representam os agentes físicos, químicos e biológicos trazidos do inventário do PGR. Itens marcados como pendentes mantêm o LTCAT em condição preliminar.",
    ),
  ];

  if (!rows.length) {
    children.push(
      paragraph(
        "Nenhum agente físico, químico ou biológico foi selecionado no inventário. Essa ausência deve ser confirmada em inspeção antes de qualquer conclusão previdenciária.",
      ),
    );
  }

  rows.forEach((row, index) => {
    const assessment = row.assessment;
    const fill = ltcatRowIsComplete(row) ? COLORS.green : COLORS.yellow;
    children.push(
      new Paragraph({ children: [new PageBreak()] }),
      subheading(`6.${index + 1} ${row.roleName} – ${row.agentType}`),
      table(
        ["Elemento", "Registro técnico"],
        [
          ["GHE e função", `${row.ghe} – ${row.roleName}`],
          ["Agente e fonte", `${row.agentType}: ${row.source}`],
          ["Possíveis agravos", row.damages],
          ["Método", assessment.method],
          ["Resultado", assessment.method === "Quantitativa" ? `${text(assessment.result)} ${text(assessment.unit, "")}`.trim() : text(assessment.result, "Avaliação qualitativa ou pendente")],
          ["Critério", text(assessment.criterion)],
          ["Circunstâncias da exposição", text(assessment.exposure)],
          ["Proteção coletiva", text(assessment.epc)],
          ["Proteção individual", text(assessment.epi)],
          ["Eficácia dos controles", assessment.effectiveness],
          ["Conclusão", assessment.conclusion],
        ],
        [31, 69],
      ),
      new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        layout: TableLayoutType.FIXED,
        borders,
        rows: [
          new TableRow({
            cantSplit: true,
            children: [cell(ltcatRowIsComplete(row) ? "Registro preenchido para revisão e assinatura técnica" : "Pendente: completar avaliação antes do uso previdenciário", { fill, bold: true })],
          }),
        ],
      }),
      gap(),
    );
  });

  children.push(
    new Paragraph({ children: [new PageBreak()] }),
    heading("7. Proteções e eficácia"),
    paragraph(
      "A simples indicação de EPC ou EPI não prova a eliminação ou neutralização da nocividade. O responsável técnico deve verificar seleção, adequação ao agente, manutenção, registros, uso efetivo e demais requisitos aplicáveis antes de concluir sobre a eficácia.",
    ),
    heading("8. Conclusão técnica"),
    pending.length
      ? paragraph(
          `Este documento permanece preliminar porque ${pending.length} de ${rows.length || 1} avaliação(ões) ainda não possuem todos os elementos técnicos necessários. Não usar as conclusões no PPP ou para caracterizar tempo especial até a conclusão da inspeção e a assinatura do profissional habilitado.`,
        )
      : paragraph(
          "As conclusões por agente constam nas respectivas fichas. Elas devem ser confirmadas pelo profissional habilitado, que responde pela metodologia, pelos dados, pelo enquadramento e pela compatibilidade com o ambiente avaliado.",
        ),
    heading("9. Atualização e relação com o PPP"),
    paragraph(
      "O empregador deve manter o laudo atualizado sempre que houver mudança de processo, ambiente, agentes, intensidade, jornada, proteção coletiva ou individual ou qualquer condição capaz de alterar a exposição. As informações do PPP devem permanecer coerentes com o LTCAT válido para o período correspondente.",
    ),
    heading("10. Responsabilidade técnica"),
    paragraph(
      "Declaro que as informações, avaliações e conclusões deste LTCAT correspondem às condições verificadas e aos critérios vigentes na data indicada.",
    ),
    new Paragraph({ spacing: { before: 620 } }),
    table(
      ["Profissional habilitado", "Habilitação e registro", "Data e assinatura"],
      [[text(state.ltcat.professionalName), `${text(state.ltcat.professionalTitle)} · ${text(state.ltcat.professionalRegistration)}`, "________________________________"]],
      [40, 30, 30],
    ),
    gap(),
    paragraph(
      "Este arquivo é uma estrutura técnica editável. O responsável deve conferir a legislação previdenciária vigente, os critérios de avaliação e todos os dados antes da assinatura.",
      { italic: true },
    ),
  );

  return baseDocument(
    `LTCAT – ${text(state.condoName)}`,
    "Laudo Técnico das Condições Ambientais do Trabalho",
    state,
    "LTCAT",
    children,
  );
}

function filename(value: string) {
  return value
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-zA-Z0-9]+/g, "-")
    .replace(/^-|-$/g, "")
    .toLowerCase() || "condominio";
}

function triggerDownload(blob: Blob, name: string) {
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = name;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  window.setTimeout(() => URL.revokeObjectURL(url), 1_000);
}

export function makePcmsoFilename(state: PgrState) {
  return `PCMSO-${filename(text(state.condoName, "condominio"))}.docx`;
}

export function makeLtcatFilename(state: PgrState) {
  return `LTCAT-${filename(text(state.condoName, "condominio"))}.docx`;
}

export async function downloadPcmsoDocx(state: PgrState) {
  triggerDownload(await Packer.toBlob(buildPcmsoDocument(state)), makePcmsoFilename(state));
}

export async function downloadLtcatDocx(state: PgrState) {
  triggerDownload(await Packer.toBlob(buildLtcatDocument(state)), makeLtcatFilename(state));
}
