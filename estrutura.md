
# 🗂️ Estrutura da Base de Dados — Plataforma de Cursos

> Atualizado automaticamente via script Apps Script  
> Cada guia (aba) representa uma entidade do sistema.  
> A segunda coluna exibe o formato JSON das colunas dessa guia.

---

## **Estrutura**
```json
{"Guia":"","Estrutura (JSON)":""}
```

## **CURIOSIDADES**
```json
{"A planilha “padrão” tem 1.048.576 linhas × 16.384 colunas (até a coluna XFD).":""}
```

## **UserProfiles**
```json
{"userId":"","phone":"","role":"","bio":"","photoFileId":"","photoUrl":"","folderId":"","updatedAt":""}
```

## **PasswordResets**
```json
{"userId":"","email":"","codeHash":"","createdAt":"","expiresAt":"","usedAt":"","lastSentAt":"","attempts":""}
```

## **QuestionBank**
```json
{"id":"","courseId":"","cohortId":"","moduleId":"","lessonId":"","question":"","questionType":"","optionsJSON":"","correctAnswer":"","points":"","explanation":"","status":"","createdBy":"","createdAt":"","updatedAt":""}
```
