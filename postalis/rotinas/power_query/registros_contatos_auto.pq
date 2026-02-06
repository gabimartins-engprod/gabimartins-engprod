// =====================================================================================
// CONSULTA: Registros_Contatos_Auto
// SISTEMA: Rotina_Dados – Postalis
// FINALIDADE: A partir da base Correção_Automática, extrair Campos de Plano, Nome e
//             Matrícula a partir do campo Task, mantendo ID, Parent task ID e Created on.
// ORIGEM DOS DADOS: Correção_Automática
// DESTINO: Base intermediária para Registros_Contatos_Final
// VERSÃO: v1.1 – atualizado em 05/02/2026
// RESPONSÁVEL: Gabi (Postalis)
// OBSERVAÇÕES IMPORTANTES:
// - Identifica automaticamente possíveis nomes de pessoa no texto.
// - Extrai matrícula a partir de padrões numéricos (6 a 12 dígitos) e retorna apenas números.
// - Inclui Parent task ID na saída.
// - Remove duplicidades por ID mantendo o registro mais recente (Created on desc).
// =====================================================================================
let
    FonteFilt=Table.SelectColumns(Correção_Automática,{"ID","Parent task ID","Task","Created on"},MissingField.Ignore),
    Tipos0=Table.TransformColumnTypes(FonteFilt,{{"ID",type text},{"Parent task ID",type text},{"Task",type text},{"Created on",type datetime}},"pt-BR"),
    NormalizaSep=(t as nullable text) as text=>let s0=if t=null then "" else Text.From(t),s1=Text.Replace(Text.Replace(Text.Replace(s0,"—","-"),"–","-")," - ","-"),s2=Text.Replace(Text.Replace(s1,"  -","-"),"-  ","-"),s3=Text.Trim(s2) in s3,
    ExtraiPlano=(t as nullable text) as nullable text=>let s=NormalizaSep(t),partes=List.Transform(Text.Split(s,"-"),each Text.Trim(_)),plano=if List.Count(partes)>0 then Text.Upper(partes{0}) else null in plano,
    PareceNome=(p as text) as logical=>let u=Text.Upper(Text.Trim(p)),temEspaco=Text.Contains(u," "),temDig=List.AnyTrue(List.Transform(Text.ToList(u),each _>="0" and _<="9")),curto=Text.Length(u)<4,ban={"PENSAO","PENSÃO","PECULIO","PECÚLIO","POSTALPREV","NOVO","REGULAMENTO","PLANO","DEFINIR","CONTROLAR","ANALISAR","ENCERRAR","ALTERACAO","ALTERAÇÃO","SUBTAREFA","INCLUSAO","INCLUSÃO","SALDADO","BD","ÚNICO","UNICO","FAMILIAR","COMPROVAÇÃO","COMPROVACAO"},tokenBan=List.AnyTrue(List.Transform(ban,each Text.Contains(u,_))) in temEspaco and (not temDig) and (not curto) and (not tokenBan),
    ExtraiNome=(t as nullable text) as nullable text=>let s=NormalizaSep(t),ps=List.Transform(Text.Split(s,"-"),each Text.Trim(_)),cand=List.Select(ps,each PareceNome(_)),nome=if List.Count(cand)>0 then Text.Proper(cand{0}) else null in nome,
    ExtraiMatricula=(t as nullable text) as nullable text=>let s0=if t=null then "" else Text.From(t),semObs=Text.BeforeDelimiter(s0,"(",1),toks=Text.SplitAny(semObs," "&Character.FromNumber(9)&Character.FromNumber(160)),nums=List.Transform(toks,each Text.Select(_,{"0".."9"})),ok=List.Transform(nums,each let n=Text.Length(_) in n>=6 and n<=12),idx=List.PositionOf(ok,true),out=if idx<>-1 then nums{idx} else null in out,
    AddPlano=Table.AddColumn(Tipos0,"Plano",each ExtraiPlano([Task]),type text),
    AddNome=Table.AddColumn(AddPlano,"Nome",each ExtraiNome([Task]),type text),
    AddMat=Table.AddColumn(AddNome,"Matrícula",each ExtraiMatricula([Task]),type text),
    Sele=Table.SelectColumns(AddMat,{"ID","Parent task ID","Created on","Plano","Nome","Matrícula"},MissingField.Ignore),
    Ordenado=Table.Sort(Sele,{{"Created on",Order.Descending}}),
    Uniq=Table.Distinct(Ordenado,{"ID"}),
    Final=Table.Sort(Uniq,{{"ID",Order.Ascending}})
in
    Final
