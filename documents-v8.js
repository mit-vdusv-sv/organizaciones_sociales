(function () {
  "use strict";

  const {
    AlignmentType,
    BorderStyle,
    Document,
    Footer,
    HeadingLevel,
    Packer,
    PageNumber,
    Paragraph,
    Table,
    TableCell,
    TableRow,
    TextRun,
    WidthType,
  } = window.docx;

  const AGREEMENT = "Acuerdo Ministerial Nro. MIT-MIT-25-60-ACU de 8 de noviembre de 2025";

  function dateLong(value) {
    if (!value) return "[FECHA]";
    const [year, month, day] = value.split("-").map(Number);
    return new Intl.DateTimeFormat("es-EC", { day: "numeric", month: "long", year: "numeric", timeZone: "UTC" }).format(new Date(Date.UTC(year, month - 1, day)));
  }

  function safe(value, fallback = "[POR COMPLETAR]") {
    return String(value || "").trim() || fallback;
  }

  function upper(value) {
    return safe(value).toLocaleUpperCase("es-EC");
  }

  function locationText(data) {
    return `domiciliada en la parroquia ${safe(data.parish)}, cantón ${safe(data.canton)}, provincia de ${safe(data.province)}`;
  }

  function organizationText(data) {
    return `${upper(data.organizationName)}, ${locationText(data)}`;
  }

  function run(text, options = {}) {
    return new TextRun({ text, font: "Times New Roman", size: 21, ...options });
  }

  function para(text, options = {}) {
    return new Paragraph({
      alignment: options.alignment || AlignmentType.JUSTIFIED,
      spacing: { after: options.after ?? 140, line: options.line ?? 276 },
      indent: options.indent,
      keepNext: options.keepNext,
      children: [run(text, { bold: options.bold, italics: options.italics, size: options.size || 21 })],
    });
  }

  function mixedPara(parts, options = {}) {
    return new Paragraph({
      alignment: options.alignment || AlignmentType.JUSTIFIED,
      spacing: { after: options.after ?? 140, line: options.line ?? 276 },
      children: parts.map((part) => run(part.text, part)),
    });
  }

  function title(text) {
    return new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 220 },
      keepNext: true,
      children: [run(text, { bold: true, size: 24 })],
    });
  }

  function sectionHeading(text) {
    return new Paragraph({
      heading: HeadingLevel.HEADING_1,
      spacing: { before: 200, after: 120 },
      keepNext: true,
      children: [run(text, { bold: true, size: 21 })],
    });
  }

  function centered(text, options = {}) {
    return para(text, { ...options, alignment: AlignmentType.CENTER });
  }

  function numberedItem(number, text) {
    return new Paragraph({
      alignment: AlignmentType.JUSTIFIED,
      spacing: { after: 90, line: 260 },
      indent: { left: 360, hanging: 300 },
      children: [run(`${number}. `, { bold: true }), run(text)],
    });
  }

  function bullet(text) {
    return new Paragraph({
      alignment: AlignmentType.JUSTIFIED,
      spacing: { after: 90, line: 260 },
      bullet: { level: 0 },
      children: [run(text)],
    });
  }

  function signature(data, signerName, signerRole) {
    return [
      para("", { after: 520 }),
      centered(upper(signerName), { bold: true, after: 40 }),
      centered(upper(signerRole), { bold: true, after: 40 }),
      centered("MINISTERIO DE INFRAESTRUCTURA Y TECNOLOGÍA", { bold: true }),
    ];
  }

  function pageFooter() {
    return new Footer({
      children: [
        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [run("Página ", { size: 17, color: "687783" }), new TextRun({ children: [PageNumber.CURRENT], font: "Times New Roman", size: 17, color: "687783" })],
        }),
      ],
    });
  }

  function makeDocument(children) {
    return new Document({
      creator: "Ministerio de Infraestructura y Tecnología",
      title: "Trámite de personalidad jurídica",
      description: `Documento generado conforme al ${AGREEMENT}`,
      styles: {
        default: { document: { run: { font: "Times New Roman", size: 21 }, paragraph: { spacing: { line: 276 } } } },
        paragraphStyles: [
          { id: "Title", name: "Title", basedOn: "Normal", next: "Normal", run: { font: "Times New Roman", size: 24, bold: true, color: "000000" }, paragraph: { alignment: AlignmentType.CENTER, spacing: { after: 220 } } },
          { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", run: { font: "Times New Roman", size: 21, bold: true, color: "000000" }, paragraph: { spacing: { before: 200, after: 120 }, keepNext: true } },
        ],
      },
      sections: [{
        properties: {
          page: { margin: { top: 1134, right: 1134, bottom: 1134, left: 1418 } },
        },
        footers: { default: pageFooter() },
        children,
      }],
    });
  }

  const legalBasis = [
    "El numeral 13 del artículo 66 y el artículo 96 de la Constitución de la República reconocen el derecho de asociación y las formas de organización de la sociedad.",
    "Los artículos 226 y 227 de la Constitución sujetan la actuación administrativa a las competencias atribuidas por el ordenamiento jurídico y a los principios que rigen la administración pública.",
    "Los artículos 565 y 567 del Código Civil regulan el establecimiento de las personas jurídicas y la aprobación de los estatutos de las corporaciones.",
    "Los artículos 3, 4, 5, 47 y 69 del Código Orgánico Administrativo regulan los principios de eficacia, eficiencia y calidad, la representación de las administraciones públicas y la delegación de competencias.",
    "Los artículos 30 y 36 de la Ley Orgánica de Participación Ciudadana regulan las formas de organización social y la obtención de personalidad jurídica.",
    `El ${AGREEMENT} expidió el Reglamento para la Gestión y Delegación de los Trámites Administrativos para la Vida Jurídica de las Organizaciones Sociales Sin Fines de Lucro del Ministerio de Infraestructura y Tecnología.`,
    "El artículo 14 del referido Reglamento establece los requisitos para el otorgamiento de personalidad jurídica, incluidos la solicitud, el acta constitutiva, el estatuto, la nómina de la directiva provisional, el domicilio legal y la declaración juramentada de patrimonio inicial.",
    "El artículo 16 determina el procedimiento de revisión, informe, aprobación, subsanación y archivo del trámite.",
    "El artículo 17 dispone que la directiva provisional tendrá una vigencia de treinta días y que, si la organización no elige su directiva estatutaria dentro de ese plazo, quedará en acefalía.",
    "La Disposición General Tercera ordena que los trámites presentados hasta la fecha de suscripción del Acuerdo que no hubieren sido atendidos o estuvieren en proceso se tramiten conforme al nuevo Reglamento.",
  ];

  const resolutionLegalBasis = [
    "el numeral 13 del artículo 66 de la Constitución de la República del Ecuador reconoce y garantiza el derecho de las personas a asociarse, reunirse y manifestarse de forma libre y voluntaria;",
    "el artículo 96 de la Constitución de la República del Ecuador reconoce todas las formas de organización de la sociedad y dispone que garanticen la democracia interna, la alternabilidad de sus dirigentes y la rendición de cuentas;",
    "el numeral 1 del artículo 154 de la Constitución de la República del Ecuador atribuye a las ministras y ministros de Estado la rectoría de las políticas públicas de su área y la expedición de los acuerdos y resoluciones administrativas que requiera su gestión;",
    "el artículo 226 de la Constitución de la República del Ecuador dispone que las instituciones, organismos y servidores públicos ejercerán únicamente las competencias y facultades atribuidas en la Constitución y la ley;",
    "el artículo 227 de la Constitución de la República del Ecuador establece que la administración pública constituye un servicio a la colectividad y se rige, entre otros, por los principios de eficacia, eficiencia, calidad, coordinación, participación, planificación y transparencia;",
    "el artículo 565 del Código Civil determina que no son personas jurídicas las fundaciones o corporaciones que no se hayan establecido en virtud de una ley o que no hayan sido aprobadas por la autoridad competente;",
    "el artículo 567 del Código Civil dispone que los estatutos de las corporaciones serán sometidos a aprobación y que esta procederá cuando no contengan disposiciones contrarias al orden público, a las leyes o a las buenas costumbres;",
    "los artículos 3, 4 y 5 del Código Orgánico Administrativo consagran los principios de eficacia, eficiencia y calidad que orientan las actuaciones de las administraciones públicas;",
    "el artículo 7 del Código Orgánico Administrativo establece el principio de desconcentración y privilegia la distribución objetiva y la delegación de funciones para acercar la administración a las personas;",
    "el artículo 47 del Código Orgánico Administrativo dispone que la máxima autoridad administrativa ejerce la representación legal de la entidad pública en los actos y relaciones jurídicas sujetos a su competencia;",
    "los artículos 69 y 70 del Código Orgánico Administrativo regulan la delegación de competencias, su contenido, alcance, responsabilidad y publicidad;",
    "los artículos 30 y 36 de la Ley Orgánica de Participación Ciudadana reconocen las formas de organización social sin fines de lucro y establecen que su constitución con personalidad jurídica se perfecciona mediante la aprobación y registro del estatuto por la autoridad competente;",
    "los artículos 1 y 3 de la Ley Orgánica para la Optimización y Eficiencia de Trámites Administrativos establecen la simplificación de los trámites, la reducción de cargas y el control posterior de la información presentada por los administrados;",
    "los artículos 17 y 55 del Estatuto del Régimen Jurídico y Administrativo de la Función Ejecutiva reconocen la competencia de los ministros para el despacho de los asuntos de sus carteras de Estado y permiten delegar atribuciones en autoridades u órganos de inferior jerarquía;",
    "la Norma de Control Interno 200-05 de la Contraloría General del Estado dispone que la delegación debe acompañarse de la autoridad necesaria y que el delegado responde por las decisiones y omisiones adoptadas en su ejercicio;",
    "mediante Decreto Ejecutivo Nro. 102 de 15 de agosto de 2025 se dispuso la fusión institucional y la conformación del ministerio rector en materia de infraestructura;",
    "la Ley Orgánica de Transparencia Social estableció el marco jurídico aplicable a las organizaciones sociales sin fines de lucro y fortaleció las reglas de transparencia, integridad, registro, control y supervisión;",
    "el Reglamento General a la Ley Orgánica de Transparencia Social desarrolló los requisitos, procedimientos, autoridades competentes y mecanismos aplicables a la constitución, funcionamiento, reorganización y disolución de las organizaciones sociales sin fines de lucro;",
    `el ${AGREEMENT} expidió el Reglamento para la Gestión y Delegación de los Trámites Administrativos para la Vida Jurídica de las Organizaciones Sociales Sin Fines de Lucro del Ministerio de Infraestructura y Tecnología;`,
    "los artículos 1 y 2 del referido Reglamento determinan su objeto y ámbito nacional de aplicación para el otorgamiento de personalidad jurídica y los demás actos relacionados con la vida jurídica de las organizaciones sociales bajo competencia de esta Cartera de Estado;",
    "el artículo 5 del Reglamento determina los ámbitos materiales en los cuales el Ministerio ejerce competencia para otorgar personalidad jurídica y emitir actos administrativos relacionados con las organizaciones sociales;",
    "los artículos 6 y 7 del Reglamento distribuyen y delegan la competencia entre las subsecretarías y direcciones distritales, y les atribuyen la aprobación de informes, estatutos, personalidad jurídica, directivas y demás actos de la vida jurídica de las organizaciones sociales;",
    "los artículos 8, 9, 10, 11 y 12 del Reglamento regulan las asambleas, el contenido mínimo de las convocatorias, la verificación del quórum y la responsabilidad de los solicitantes y de las certificaciones;",
    "el artículo 14 del Reglamento establece los documentos y contenidos requeridos para el otorgamiento de personalidad jurídica, incluidos la solicitud, el acta constitutiva, el estatuto, la directiva provisional, el domicilio legal y la declaración juramentada de patrimonio inicial;",
    "el artículo 15 del Reglamento exige que la denominación de la organización sea claramente distinguible de otras organizaciones activas registradas;",
    "el artículo 16 del Reglamento establece el procedimiento de revisión, emisión del informe motivado, aprobación, subsanación y archivo de la solicitud;",
    "el artículo 17 del Reglamento dispone que la directiva provisional tendrá una vigencia de treinta días y que la organización deberá elegir su directiva estatutaria dentro de dicho plazo;",
    "la Disposición General Tercera del Reglamento ordena que los trámites presentados hasta la fecha de suscripción del Acuerdo que no hubieren sido atendidos o se encontraren en proceso sean tramitados conforme al Reglamento vigente;",
    "la Disposición Derogatoria del Acuerdo dejó sin efecto los instrumentos ministeriales anteriores que regulaban estos trámites en el Ministerio de Desarrollo Urbano y Vivienda y en el Ministerio de Transporte y Obras Públicas;",
  ];

  function foundersTable(members) {
    const border = { style: BorderStyle.SINGLE, size: 2, color: "808080" };
    const borders = { top: border, bottom: border, left: border, right: border, insideHorizontal: border, insideVertical: border };
    const header = new TableRow({
      tableHeader: true,
      children: ["N.º", "NOMBRES Y APELLIDOS", "CÉDULA", "NACIONALIDAD"].map((text) => new TableCell({
        shading: { fill: "DCE7EC" },
        children: [centered(text, { bold: true, size: 18, after: 0 })],
      })),
    });
    const rows = (members || []).map((member, index) => new TableRow({
      children: [String(index + 1), upper(member.name), safe(member.id), upper(member.nationality)].map((text, col) => new TableCell({
        children: [para(text, { alignment: col === 1 ? AlignmentType.LEFT : AlignmentType.CENTER, size: 18, after: 0, line: 220 })],
      })),
    }));
    return new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      columnWidths: [700, 4800, 1900, 2200],
      borders,
      rows: [header, ...rows],
    });
  }

  function transitionParagraph(data) {
    if (!data.originalCaseDate || data.originalCaseDate > "2025-11-08") return [];
    return [para(`El expediente registra como fecha original de presentación el ${dateLong(data.originalCaseDate)}. En aplicación de la Disposición General Tercera del ${AGREEMENT}, su sustanciación continúa conforme al Reglamento vigente.`)];
  }

  function buildResolution(data) {
    const org = organizationText(data);
    const children = [
      centered(`RESOLUCIÓN Nro. ${safe(data.resolutionNumber)}`, { bold: true, size: 23 }),
      centered(upper(data.authorityName), { bold: true, after: 40 }),
      centered(upper(data.authorityRole), { bold: true }),
      centered(upper(data.administrativeUnit), { bold: true, after: 220 }),
      centered("CONSIDERANDO", { bold: true, after: 180 }),
      ...resolutionLegalBasis.map((text) => mixedPara([{ text: "Que, ", bold: true }, { text }])),
      mixedPara([{ text: "Que, ", bold: true }, { text: `mediante ${safe(data.appointmentNumber)}, de ${dateLong(data.appointmentDate)}, se designó a ${safe(data.authorityName)} como ${safe(data.authorityRole)} de ${safe(data.administrativeUnit)};` }]),
      mixedPara([{ text: "Que, ", bold: true }, { text: `mediante ${data.requestLetterNumber ? `oficio Nro. ${data.requestLetterNumber}` : "oficio sin número"}, de ${dateLong(data.requestLetterDate)}, ingresado con trámite Nro. ${safe(data.caseNumber)}, de ${dateLong(data.entryDate)}, ${safe(data.applicantName)}, en calidad de ${safe(data.applicantRole)}, solicitó la aprobación y registro del estatuto y el otorgamiento de personalidad jurídica de ${org};` }]),
      mixedPara([{ text: "Que, ", bold: true }, { text: `los miembros fundadores aprobaron el estatuto en asamblea general constitutiva de ${dateLong(data.assemblyDate)};` }]),
      mixedPara([{ text: "Que, ", bold: true }, { text: `mediante informe jurídico Nro. ${safe(data.legalReportNumber)}, de ${dateLong(data.legalReportDate)}, ${safe(data.legalAnalystName)}, ${safe(data.legalAnalystRole)}, concluyó que la organización cumple los requisitos del artículo 14 del Reglamento y que procede el otorgamiento de personalidad jurídica; y,` }]),
      ...transitionParagraph(data),
      para(`En ejercicio de las atribuciones delegadas mediante el ${AGREEMENT},`),
      centered("RESUELVE", { bold: true, size: 23, after: 180 }),
      mixedPara([{ text: "Artículo 1.- ", bold: true }, { text: `Aprobar y registrar el estatuto de ${org}, ${data.statuteModification === "con" ? "con las modificaciones incorporadas" : "sin modificación"}.` }]),
      mixedPara([{ text: "Artículo 2.- ", bold: true }, { text: `Otorgar personalidad jurídica a ${org} y registrar como miembros fundadores a las siguientes personas:` }]),
      foundersTable(data.members),
      para("", { after: 50 }),
      mixedPara([{ text: "Artículo 3.- ", bold: true }, { text: `Disponer que ${upper(data.organizationName)}, dentro del plazo máximo de treinta días contado desde el otorgamiento de personalidad jurídica, elija su directiva estatutaria y solicite su registro ante ${safe(data.administrativeUnit)}, conforme al artículo 17 del Reglamento.` }]),
      mixedPara([{ text: "Artículo 4.- ", bold: true }, { text: `La veracidad y autenticidad de los documentos presentados, así como de los objetivos y fines declarados por ${upper(data.organizationName)}, son de responsabilidad de los peticionarios.` }]),
      mixedPara([{ text: "Artículo 5.- ", bold: true }, { text: `La organización deberá desarrollar sus actividades dentro de los fines y objetivos establecidos en su estatuto y cumplir las obligaciones previstas en el ordenamiento jurídico.` }]),
      mixedPara([{ text: "Artículo 6.- ", bold: true }, { text: "La presente resolución no autoriza la realización de actividades con fines de lucro ni actividades comerciales ajenas a los fines estatutarios." }]),
      mixedPara([{ text: "Artículo 7.- ", bold: true }, { text: "Los conflictos internos se resolverán conforme al estatuto, a los mecanismos alternativos de solución de conflictos y a las acciones previstas en el ordenamiento jurídico, con respeto al debido proceso." }]),
      sectionHeading("DISPOSICIONES FINALES"),
      mixedPara([{ text: "PRIMERA.- ", bold: true }, { text: `Notifíquese el contenido de esta resolución a ${upper(data.organizationName)}.` }]),
      mixedPara([{ text: "SEGUNDA.- ", bold: true }, { text: `De la ejecución de la presente resolución encárguese ${safe(data.administrativeUnit)} del Ministerio de Infraestructura y Tecnología.` }]),
      mixedPara([{ text: "TERCERA.- ", bold: true }, { text: "La presente resolución entrará en vigencia a partir de su suscripción." }]),
      para(`Dado en ${safe(data.issueCity)}, a ${dateLong(data.issueDate)}.`),
      ...signature(data, data.authorityName, data.authorityRole),
    ];
    return makeDocument(children);
  }

  function buildResponse(data) {
    const org = organizationText(data);
    const children = [
      mixedPara([{ text: "ASUNTO: ", bold: true }, { text: `RESPUESTA A LA SOLICITUD DE OTORGAMIENTO DE PERSONALIDAD JURÍDICA DE ${upper(data.organizationName)}` }], { alignment: AlignmentType.LEFT, after: 240 }),
      para(`En atención a ${data.requestLetterNumber ? `oficio Nro. ${data.requestLetterNumber}` : "oficio sin número"}, de ${dateLong(data.requestLetterDate)}, ingresado con trámite Nro. ${safe(data.caseNumber)}, de ${dateLong(data.entryDate)}, mediante el cual ${safe(data.applicantName)}, en calidad de ${safe(data.applicantRole)}, solicitó la aprobación y registro del estatuto y el otorgamiento de personalidad jurídica de ${org}, comunico lo siguiente:`),
      para(`Una vez revisada la documentación, se verificó el cumplimiento de los requisitos establecidos en el artículo 14 del Reglamento expedido mediante el ${AGREEMENT}. Conforme a los artículos 6 y 7 del mismo instrumento, esta autoridad aprobó y registró el estatuto y otorgó personalidad jurídica a la organización mediante Resolución Nro. ${safe(data.resolutionNumber)}.`),
      ...transitionParagraph(data),
      para("La organización deberá elegir y solicitar el registro de su directiva estatutaria dentro del plazo máximo de treinta días, conforme al artículo 17 del Reglamento."),
      para("La veracidad y autenticidad de la documentación entregada al Ministerio de Infraestructura y Tecnología es responsabilidad de los miembros y del solicitante. La autoridad competente podrá ejercer el control posterior y adoptar las medidas previstas en el ordenamiento jurídico si comprueba información falsa, incompleta o adulterada."),
      para("Lo que comunico para los fines pertinentes."),
      ...signature(data, data.authorityName, data.authorityRole),
    ];
    return makeDocument(children);
  }

  function buildReport(data) {
    const org = organizationText(data);
    const children = [
      title(`INFORME JURÍDICO PARA LA APROBACIÓN Y REGISTRO DEL ESTATUTO Y EL OTORGAMIENTO DE PERSONALIDAD JURÍDICA DE ${upper(data.organizationName)}`),
      para(`Se emite el presente informe jurídico respecto de la solicitud de aprobación y registro del estatuto y otorgamiento de personalidad jurídica de ${org}.`),
      sectionHeading("I. ANTECEDENTES"),
      para(`Mediante ${data.requestLetterNumber ? `oficio Nro. ${data.requestLetterNumber}` : "oficio sin número"}, de ${dateLong(data.requestLetterDate)}, ingresado con trámite Nro. ${safe(data.caseNumber)}, de ${dateLong(data.entryDate)}, ${safe(data.applicantName)}, en calidad de ${safe(data.applicantRole)}, solicitó la aprobación y registro del estatuto y el otorgamiento de personalidad jurídica de ${org}.`),
      para(`La asamblea general constitutiva se celebró el ${dateLong(data.assemblyDate)} y aprobó la constitución de la organización, su estatuto y la directiva provisional.`),
      ...(data.additionalBackground ? [para(data.additionalBackground)] : []),
      ...transitionParagraph(data),
      sectionHeading("II. NORMATIVA JURÍDICA APLICABLE"),
      ...legalBasis.map((text, index) => numberedItem(index + 1, text)),
      sectionHeading("III. VERIFICACIÓN DE REQUISITOS"),
      para("De la revisión del expediente se verificó el cumplimiento de los requisitos previstos en el artículo 14 del Reglamento, conforme al siguiente detalle:"),
      ...(data.checklist || []).filter((item) => item.status === "cumple" || item.status === "no_aplica").map((item) => bullet(`${item.title}: ${item.status === "no_aplica" ? "No aplica al trámite" : "Cumple"}.`)),
      sectionHeading("IV. ANÁLISIS JURÍDICO"),
      para(`La documentación presentada acredita la voluntad de constituir ${upper(data.organizationName)}, identifica a sus miembros fundadores y directiva provisional, determina su domicilio y alcance territorial, incorpora la declaración juramentada de patrimonio inicial y contiene un estatuto que regula los elementos mínimos exigidos por el artículo 14.`),
      para(`Los fines declarados se relacionan con ${safe(data.actionArea)} y se encuentran dentro del ámbito de competencia del Ministerio de Infraestructura y Tecnología. El estatuto no contiene disposiciones contrarias al orden público ni a la normativa aplicable.`),
      sectionHeading("V. CONCLUSIÓN"),
      para(`Con base en los antecedentes, la normativa aplicable y la verificación efectuada, se concluye que ${org} cumple los requisitos para la aprobación y registro de su estatuto y el otorgamiento de personalidad jurídica. En consecuencia, procede la emisión de la resolución correspondiente ${data.statuteModification === "con" ? "con las modificaciones incorporadas" : "sin modificación"}.`),
      ...signature(data, data.legalAnalystName, data.legalAnalystRole),
    ];
    return makeDocument(children);
  }

  function buildObservation(data) {
    const failures = (data.checklist || []).filter((item) => item.status === "no_cumple" || item.status === "pendiente");
    const children = [
      mixedPara([{ text: "ASUNTO: ", bold: true }, { text: `OBSERVACIONES A LA SOLICITUD DE OTORGAMIENTO DE PERSONALIDAD JURÍDICA DE ${upper(data.organizationName)}` }], { alignment: AlignmentType.LEFT, after: 240 }),
      para(`En atención a ${data.requestLetterNumber ? `oficio Nro. ${data.requestLetterNumber}` : "oficio sin número"}, de ${dateLong(data.requestLetterDate)}, ingresado con trámite Nro. ${safe(data.caseNumber)}, de ${dateLong(data.entryDate)}, mediante el cual se solicitó el otorgamiento de personalidad jurídica de ${organizationText(data)}, comunico lo siguiente:`),
      sectionHeading("I. FUNDAMENTO JURÍDICO"),
      para(`El artículo 14 del Reglamento expedido mediante el ${AGREEMENT} establece los requisitos para el otorgamiento de personalidad jurídica. El numeral 4 del artículo 16 dispone que, cuando la documentación no satisfaga dichos requisitos, la autoridad competente otorgará a la organización un plazo máximo de veinte días para completar o subsanar y volver a presentar la documentación.`),
      ...transitionParagraph(data),
      sectionHeading("II. OBSERVACIONES"),
      ...(failures.length ? failures.map((item, index) => numberedItem(index + 1, `${item.title}. ${item.observation || (item.status === "pendiente" ? "El requisito no ha sido verificado." : "Debe completarse o corregirse la documentación correspondiente.")}`)) : [para("No se registraron requisitos incumplidos. Revise el estado del checklist antes de emitir este documento.")]),
      sectionHeading("III. SUBSANACIÓN"),
      para("La organización deberá presentar la documentación completa y corregida dentro del plazo máximo de veinte días. Una vez reingresada, la unidad competente efectuará una nueva revisión. Si persiste el incumplimiento, se dispondrá el archivo y cierre del trámite, con la notificación correspondiente."),
      para("La documentación de subsanación deberá identificar el número del trámite y responder de manera ordenada a cada observación. La veracidad y autenticidad de los documentos presentados son responsabilidad de los miembros y del solicitante."),
      para("Lo que comunico para los fines pertinentes."),
      ...signature(data, data.authorityName, data.authorityRole),
    ];
    return makeDocument(children);
  }

  async function downloadDocument(kind, data) {
    const builders = { resolution: buildResolution, response: buildResponse, report: buildReport, observation: buildObservation };
    const labels = {
      resolution: "Resolucion_otorgamiento_personalidad_juridica",
      response: "Oficio_respuesta_personalidad_juridica",
      report: "Informe_juridico_personalidad_juridica",
      observation: "Observacion_otorgamiento_personalidad_juridica",
    };
    const generatedDocument = builders[kind](data);
    const blob = await Packer.toBlob(generatedDocument);
    const url = URL.createObjectURL(blob);
    const link = window.document.createElement("a");
    link.href = url;
    const orgSlug = (data.organizationName || "organizacion").normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/[^a-zA-Z0-9]+/g, "_").replace(/^_|_$/g, "");
    link.download = `${labels[kind]}_${orgSlug}.docx`;
    window.document.body.appendChild(link);
    link.click();
    link.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  }

  window.MITDocuments = { downloadDocument, dateLong };
})();
