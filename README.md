# Office → PDF (versão 2.0, Windows)

Programa com janela: escolhes uma pasta e ele converte **em massa** para PDF:

| Tipo | Extensões |
|---|---|
| Word | `.doc` `.docx` |
| Excel | `.xls` `.xlsx` — respeita a configuração de impressão de cada ficheiro |
| PowerPoint | `.ppt` `.pptx` |

Os PDF ficam numa subpasta `PDF` dentro da pasta escolhida (com opção de incluir subpastas).
Tudo corre no computador de quem usa — nada vai para a nuvem.

- **Motor:** usa o Microsoft Office instalado (Word, Excel, PowerPoint) e, se faltar um deles
  ou falhar num ficheiro, passa para o **LibreOffice** (gratuito). Também é possível forçar um dos dois.
- **Nomes repetidos:** `folha.xls` e `folha.xlsx` na mesma pasta dão `folha (xls).pdf` e `folha (xlsx).pdf`.
- **Macros** ficam desativadas durante a conversão. Ficheiros temporários (`~$...`) são ignorados.

## Para quem vai usar

1. Descarregar `Office_para_PDF.exe` da página **Releases** do repositório.
2. Duplo clique. Se o Windows avisar (o ficheiro não está assinado): **Mais informações → Executar mesmo assim**.
3. **Escolher pasta…**, marcar os tipos pretendidos e **Converter para PDF**.

Requisito: Microsoft Office **ou** LibreOffice instalado. Não é preciso Python.
Linha de comandos: `Office_para_PDF.exe "C:\pasta" --subpastas --tipos=word,excel --motor=auto`

## Atualizar o repositório que já tens (versão 1 → 2)

1. No repositório, **Add file → Upload files** e carrega: `office_para_pdf.py`, `requirements.txt`
   e a pasta `tests` (com `criar_amostras.py`). Carrega em **Commit changes**.
2. Abre `.github/workflows/build-exe.yml` → lápis (**Edit**) → apaga o conteúdo todo, cola o
   conteúdo do novo `build-exe.yml` → **Commit changes**.
3. Abre o antigo `word_para_pdf.py` → menu **⋯ → Delete file** → **Commit changes** (deixou de ser usado).
4. **Actions → Gerar .exe → Run workflow**. Espera pelo visto verde (3 a 6 minutos). O passo
   «Teste de conversão real (LibreOffice)» é informativo: se falhar, não impede o resto.
5. **Releases → Draft a new release** → etiqueta `v2.0.0` → **Publish release**.
   Passado pouco tempo, o `Office_para_PDF.exe` aparece em **Assets**.

A Release `v1.0.1` (só Word) pode ficar como histórico ou ser apagada.

## O que foi testado e o que não foi

Testado (Linux, LibreOffice): conversão dos 6 tipos, subpastas, nomes repetidos, ficheiros `~$`,
configuração de impressão do Excel (folha ajustada a 1 página → 1 página; sem ajuste → 3),
janela, linha de comandos, e a lógica do Office com módulos simulados.

**Não testado:** chamadas reais ao Excel e ao PowerPoint (não havia Office disponível).
O caminho do Word reutiliza o da versão 1, que já funcionou. Se algum ficheiro falhar
com o Office, o modo «Automático» tenta o LibreOffice e o registo mostra o motivo.

## Limitações

- Ficheiros com palavra-passe: com o motor Office pode aparecer um pedido de palavra-passe
  (a conversão fica à espera). Remove a palavra-passe antes ou usa o motor LibreOffice, que
  regista o erro e continua.
- Excel: folhas ocultas não são impressas; um livro sem nada para imprimir dá erro nesse ficheiro.
  O Excel pode precisar de uma impressora instalada (a «Microsoft Print to PDF» do Windows serve).
- Alguns antivírus desconfiam de executáveis feitos com PyInstaller (falsos positivos).
