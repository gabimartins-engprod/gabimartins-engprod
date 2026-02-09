// =====================================================================================
// CONSULTA: Registros_Contatos_Auto
// SISTEMA: Rotina_Dados – Postalis
// FINALIDADE: A partir da base Correção_Automática, extrair Campos de Plano, Nome e
//             Matrícula a partir do campo Task, mantendo ID, Parent task ID e Created on.
// ORIGEM DOS DADOS: Correção_Automática
// DESTINO: Tabela Registros_Contatos_Auto no Excel (base intermediária p/ Registros_Contatos_Final)
// VERSÃO: v1.3 – atualizado em 09/02/2026
// RESPONSÁVEL: Gabi (Postalis)
// OBSERVAÇÕES IMPORTANTES:
// - Extrai Plano pelo primeiro bloco antes do separador "-" (após normalização).
// - Identifica Nome por heurística (tokens com espaço, sem dígitos e sem palavras bloqueadas).
// - Extrai Matrícula por padrão numérico (6 a 12 dígitos) e retorna apenas os dígitos.
// - Remove duplicidades por ID mantendo o registro mais recente (Created on desc).
// - Remove linhas com ID nulo/vazio antes da deduplicação (evita colapsar registros).
// - Compatível com versões de Power Query que NÃO suportam Text.RemoveDiacritics.
// =====================================================================================
let
    Fonte0 = Correção_Automática,

    ValidarID =
        if Table.HasColumns(Fonte0, {"ID"})
        then Fonte0
        else error "Registros_Contatos_Auto: coluna essencial 'ID' não encontrada na Correção_Automática.",

    FonteFilt =
        Table.SelectColumns(
            ValidarID,
            {"ID","Parent task ID","Task","Created on"},
            MissingField.Ignore
        ),

    Tipos0 =
        Table.TransformColumns(
            FonteFilt,
            {
                {"ID", each if _=null then null else Text.Trim(Text.From(_)), type text},
                {"Parent task ID", each if _=null then null else Text.Trim(Text.From(_)), type text},
                {"Task", each if _=null then null else Text.From(_), type text},
                {"Created on", each if _=null then null else (try DateTime.From(_) otherwise null), type datetime}
            }
        ),

    FiltraID =
        Table.SelectRows(
            Tipos0,
            each [ID] <> null and Text.Trim([ID]) <> ""
        ),

    NormalizaSep = (t as nullable text) as text =>
        let
            s0 = if t=null then "" else Text.From(t),
            s1 = Text.Replace(Text.Replace(Text.Replace(s0,"—","-"),"–","-")," - ","-"),
            s2 = Text.Replace(Text.Replace(s1,"  -","-"),"-  ","-"),
            s3 = Text.Trim(s2)
        in
            s3,

    ExtraiPlano = (t as nullable text) as nullable text =>
        let
            s = NormalizaSep(t),
            partes = List.Transform(Text.Split(s,"-"), each Text.Trim(_)),
            plano = if List.Count(partes)>0 then Text.Upper(partes{0}) else null
        in
            plano,

    RemoverAcentos = (txt as text) as text =>
        let
            mapa = {
                {"Á","A"},{"À","A"},{"Â","A"},{"Ã","A"},{"Ä","A"},
                {"É","E"},{"È","E"},{"Ê","E"},{"Ë","E"},
                {"Í","I"},{"Ì","I"},{"Î","I"},{"Ï","I"},
                {"Ó","O"},{"Ò","O"},{"Ô","O"},{"Õ","O"},{"Ö","O"},
                {"Ú","U"},{"Ù","U"},{"Û","U"},{"Ü","U"},
                {"Ç","C"}
            },
            out = List.Accumulate(mapa, txt, (state, pair) => Text.Replace(state, pair{0}, pair{1}))
        in
            out,

    NormalizaCmp = (p as nullable text) as text =>
        let
            base = if p=null then "" else Text.Upper(Text.Trim(Text.From(p))),
            semAcento = RemoverAcentos(base)
        in
            semAcento,

    PareceNome = (p as nullable text) as logical =>
        let
            u0 = NormalizaCmp(p),
            temEspaco = Text.Contains(u0," "),
            temDig = List.AnyTrue(List.Transform(Text.ToList(u0), each _>="0" and _<="9")),
            curto = Text.Length(u0) < 4,
            ban =
                {
                    "PENSAO","PECULIO","POSTALPREV","NOVO","REGULAMENTO","PLANO","DEFINIR","CONTROLAR","ANALISAR",
                    "ENCERRAR","ALTERACAO","SUBTAREFA","INCLUSAO","SALDADO","BD","UNICO","FAMILIAR",
                    "COMPROVACAO","TAREFA","PROCESSO","MIGRADO","MIGRACAO","CARTAO"
                },
            tokenBan = List.AnyTrue(List.Transform(ban, each Text.Contains(u0,_)))
        in
            temEspaco and (not temDig) and (not curto) and (not tokenBan),

    ExtraiNome = (t as nullable text) as nullable text =>
        let
            s = NormalizaSep(t),
            ps = List.Transform(Text.Split(s,"-"), each Text.Trim(_)),
            cand = List.Select(ps, each PareceNome(_)),
            nome = if List.Count(cand)>0 then Text.Proper(cand{0}) else null
        in
            nome,

    ExtraiMatricula = (t as nullable text) as nullable text =>
        let
            s0 = if t=null then "" else Text.From(t),
            semObs = Text.BeforeDelimiter(s0,"(",1),
            toks = Text.SplitAny(semObs, " " & Character.FromNumber(9) & Character.FromNumber(160)),
            nums = List.Transform(toks, each Text.Select(_, {"0".."9"})),
            ok = List.Transform(nums, each let n=Text.Length(_) in n>=6 and n<=12),
            idx = List.PositionOf(ok, true),
            out = if idx<>-1 then nums{idx} else null
        in
            out,

    AddPlano = Table.AddColumn(FiltraID, "Plano", each ExtraiPlano([Task]), type text),
    AddNome  = Table.AddColumn(AddPlano, "Nome", each ExtraiNome([Task]), type text),
    AddMat   = Table.AddColumn(AddNome, "Matrícula", each ExtraiMatricula([Task]), type text),

    Sele =
        Table.SelectColumns(
            AddMat,
            {"ID","Parent task ID","Created on","Plano","Nome","Matrícula"},
            MissingField.Ignore
        ),

    Ordenado = Table.Sort(Sele, {{"Created on", Order.Descending}}),
    Uniq = Table.Distinct(Ordenado, {"ID"}),
    Final = Table.Sort(Uniq, {{"ID", Order.Ascending}})
in
    Final
