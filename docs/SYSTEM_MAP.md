# Mapeamento da Plataforma Gamificada de Excel

Este documento explica como organizar a planilha `1DCzIOIcRBaJ3WJVQCOWg6KXyghjRGMJUHekV1fGygJ4`, quais abas utilizar, o modelo de dados e o fluxo de APIs exposto pelo Apps Script (`apps-script/Code.gs`). Ele também lista evoluções recomendadas para transformar o ambiente em uma plataforma robusta, segura e extensível.

## 1. Visão geral

A plataforma foi desenhada para entregar a jornada completa de 24 aulas (1h cada), com trilha gamificada que envolve:

- Cadastro e login de alunos/administradores;
- Distribuição de módulos e questões (objetivas e discursivas);
- Registro de respostas, cálculo automático de XP e progresso;
- Ranking dinâmico e painel administrativo com bonificações;
- Check-in diário, embeds de materiais (Excel/PPT) e geração futura de certificados.

Toda a persistência migrou do `localStorage` para a planilha Google Sheets, acessada via Apps Script publicado como Web App.

## 2. Estrutura da planilha

Crie as abas a seguir (respeitando os cabeçalhos) ou execute o script uma vez para que ele as crie automaticamente. Todas as abas usam a primeira linha como cabeçalho.

| Aba        | Cabeçalhos                                                                 | Descrição                                                                 |
|------------|----------------------------------------------------------------------------|---------------------------------------------------------------------------|
| `Users`    | `ID`, `Nome`, `Email`, `SenhaHash`, `Admin`, `XP`, `CriadoEm`, `AtualizadoEm` | Cadastro de usuários. `SenhaHash` é SHA-256 em Base64 (calculado pelo script). `Admin` aceita `TRUE`/`FALSE`. |
| `Tokens`   | `Token`, `UserID`, `ExpiresAt`, `CriadoEm`                                 | Sessões ativas. Sempre que o usuário faz login ou bootstrap autenticado um novo token é gerado (expira em 30 dias). |
| `Modules`  | `ModuleID`, `Ordem`, `Titulo`, `Descricao`, `XP`, `VideoURL`, `MaterialURL`, `Ativo` | Catálogo de módulos (24 linhas). `XP` define a recompensa máxima, `Ativo=FALSE` oculta módulos. |
| `Questions`| `ModuleID`, `QuestionID`, `Tipo`, `Enunciado`, `OpcoesJSON`, `Correta`, `Peso`, `MinCaracteres`, `Feedback` | Questões vinculadas ao módulo. `OpcoesJSON` aceita JSON ou texto separado por `||`. `Tipo` pode ser `mc` (múltipla escolha) ou `text`. |
| `Progress` | `UserID`, `ModuleID`, `Score`, `Done`, `XP`, `AnswersJSON`, `AtualizadoEm`  | Registro das tentativas. `XP` guarda a última pontuação convertida em XP, evitando duplicidade. |
| `Checkins` | `UserID`, `Date`, `XP`, `RegistradoEm`                                      | Presenças diárias e XP concedido.                                       |
| `Config`   | `Chave`, `Valor`                                                           | Configurações dinâmicas. Exemplos: `xpCheckin`, `xpPorNivel`, `ciclo`, `certificadoMinScore`. |
| `Embeds`   | `Tipo`, `URL`, `AtualizadoPor`, `AtualizadoEm`                             | Links para exibição de Excel/PPT (aceita Onedrive, SharePoint, Google Drive). |
| `AuditLog` | `Timestamp`, `Action`, `UserID`, `Payload`                                  | Trilhas para auditoria (cadastros, check-ins, atualizações).             |

> **Dica:** Utilize validação de dados na planilha (listas suspensas, etc.) para reduzir erros de digitação. As colunas `XP`, `Score`, `Peso` e `MinCaracteres` devem ser numéricas.

## 3. API exposta pelo Apps Script

Publicar o Apps Script (`Deploy → New Deployment → Web app`). Configure **Anyone** com acesso (ou restrinja para logins Google). Use o URL publicado no front-end.

Cada requisição `POST` espera um JSON `{ "action": "nomeDaAcao", "token": "...", "payload": { ... } }` e responde com `{ success, message?, ...dados }`.

### 3.1. Ações principais

| Ação              | Autenticação | Payload esperado                               | Retorno principal                                                                 |
|-------------------|--------------|------------------------------------------------|-----------------------------------------------------------------------------------|
| `bootstrap`       | opcional     | —                                              | `modules`, `cfg`, `embeds`, `ranking`, `user` (quando token válido), `adminUsers`. Revalida token. |
| `signup`          | —            | `{ nome, email, senha, admin? }`               | Novo `token`, `user`, `ranking`.                                                 |
| `login`           | —            | `{ email, senha }`                              | `token`, `user`, `ranking`, `adminUsers` se o usuário for admin.                 |
| `logout`          | obrigatório  | —                                              | Mensagem de confirmação, token revogado.                                         |
| `checkin`         | obrigatório  | —                                              | Mensagem, `user` atualizado, `ranking` e `adminUsers` (quando admin).            |
| `submitProgress`  | obrigatório  | `{ moduleId, answers: [{questionId, answer}], ... }` | Mensagem, `user` atualizado, `ranking`, `adminUsers`. XP calculado pelo backend. |
| `updateEmbed`     | admin        | `{ type: 'excel' | 'ppt', url }`               | `embeds` atualizados.                                                            |
| `awardXp`         | admin        | `{ userId, amount }`                            | `ranking`, `adminUsers` atualizados.                                             |
| `exportData`      | admin        | —                                              | Dump JSON com todas as abas para backup.                                         |
| `importData`      | admin        | `{ users?, modules?, ... }`                     | Mensagem de sucesso após sobrescrever as abas indicadas.                         |

### 3.2. Cálculo de XP e progresso

- Cada módulo possui `xp` máximo. O script confere as respostas utilizando a base da aba `Questions`.
- Perguntas podem ter pesos diferentes (`Peso`). O score percentual (`Score`) é calculado como `soma dos pesos corretos / soma dos pesos`.
- `XP` obtido = `round(xp * score%)`. O backend controla diferenças para evitar acumular XP duplicado em reenvios.
- `Done` fica `TRUE` quando o percentual ≥ 70 (ajustável pela coluna `Done` ou regra futura em `Config`).

### 3.3. Segurança e auditoria

- Senhas são armazenadas como SHA-256 Base64.
- Tokens expiram após 30 dias. Cada bootstrap autenticado gera um novo token e revoga o anterior.
- Todas as ações críticas registram um evento na aba `AuditLog`.

## 4. Ajustes no front-end (`Index`)

- `localStorage` agora guarda apenas o `token`. Todos os dados dinâmicos vêm da API.
- O painel consome as estruturas retornadas (`modules`, `ranking`, `adminUsers`, `cfg`, `embeds`).
- Ações como check-in, submissão de atividades, salvar embeds e bonificar XP usam `fetch` assíncrono para o Apps Script.
- Foi adicionada tratativa de erros com alertas amigáveis.

## 5. Roadmap de melhorias

1. **Certificados automáticos**
   - Gerar PDF via Apps Script (`DocumentApp`) quando o aluno atingir critérios (`Score` mínimo e 24 módulos concluídos).
   - Registrar em aba `Certificates` com link compartilhável.

2. **Banco de questões enriquecido**
   - Suporte a anexos (links de vídeos/arquivos). 
   - Tagging por tema (`funções`, `tabelas dinâmicas`, `Power Query`).

3. **Gamificação avançada**
   - Badges por marcos (ex.: primeiro módulo concluído, 7 check-ins consecutivos).
   - Temporizadores e livescore para desafios ao vivo.

4. **Comunicação com o aluno**
   - Envio de e-mails automáticos (GmailApp) em milestones ou quando ficar inativo.
   - Dashboard com notificações e agenda de aulas ao vivo.

5. **Analytics e BI**
   - Aba `Dashboards` usando gráficos nativos do Sheets.
   - Conector Data Studio/Looker com views pré-formatadas.

6. **Escalabilidade e segurança**
   - Migração futura para Firebase/Auth para autenticação robusta.
   - Rate limiting básico no Apps Script (checar tokens por IP/usuário).
   - Logs com `stackdriver` (BigQuery) para auditoria avançada.

7. **Integração com pagamentos**
   - Webhook de plataformas (Hotmart, Eduzz) inserindo automaticamente novos alunos na aba `Users` com status de matrícula.

8. **App mobile / PWA**
   - Aproveitar o front-end atual para publicar como Progressive Web App.
   - Implementar push notifications usando FCM/Web Push para lembretes de check-in.

9. **Suporte offline inteligente**
   - Cache dos módulos no navegador (IndexedDB) apenas para leitura, sincronizando respostas quando a conexão voltar.

10. **Trilhas alternativas e personalização**
    - Coluna `Track` nos módulos para liberar conteúdo conforme perfil (iniciante, avançado, Business Intelligence).
    - Recomendação de módulos extras baseada em desempenho.

## 6. Checklist para implantação

1. Copiar o conteúdo de `apps-script/Code.gs` para um projeto Apps Script vinculado à planilha.
2. Salvar, autorizar e publicar como Web App. Copiar a URL.
3. No front-end (`Index`), configurar `API_URL` com a URL do passo anterior e hospedar a página (GitHub Pages, Netlify, etc.).
4. Popular as abas `Modules` e `Questions` com as 24 aulas e perguntas correspondentes (pode importar CSV/JSON).
5. Ajustar a aba `Config` com valores de XP e ciclo atual.
6. Testar fluxos principais: cadastro, login, check-in, envio de atividade, ranking e painel admin.
7. Quando tudo estiver validado, habilitar novos recursos do roadmap conforme prioridade.

Com esse mapeamento a plataforma deixa de ser apenas um protótipo local e passa a utilizar a planilha como banco central, pronta para receber evoluções como certificados, automações de e-mail e integrações externas.
