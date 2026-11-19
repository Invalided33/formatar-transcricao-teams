# Formatador de transcrições do Microsoft Teams

O script lê uma transcrição exportada pelo Teams em Word (`.docx`) e cria uma cópia com uma fala por parágrafo. O horário e o nome aparecem em negrito, seguidos do texto corrido na mesma linha.

> **00:02:15 Ana Silva:** Vamos começar a reunião.

## Como usar

Com Python 3.9 ou superior, abra um terminal na pasta do projeto e informe o caminho da transcrição.

```sh
python formatar_transcricao.py "reuniao.docx"
```

A cópia será salva como `reuniao_formatado.docx` na mesma pasta da transcrição. O processamento é local e usa apenas a biblioteca padrão do Python. Para ambientes que instalam dependências a partir de um arquivo, o projeto inclui `requirements.txt`.

```sh
python -m pip install -r requirements.txt
```

## Opções

Para escolher o nome ou a pasta da cópia, use `-o`. Escolha um caminho novo, pois o programa preserva os arquivos existentes.

```sh
python formatar_transcricao.py "reuniao.docx" -o "ata formatada.docx"
```

Por padrão, o script remove todas as figuras DrawingML incorporadas ao corpo do documento, incluindo os avatares e os ícones de início e fim da transcrição. Se o documento também tiver fotos ou logos que você deseja preservar, consulte os IDs das figuras e indique somente as que deseja remover.

```sh
python formatar_transcricao.py "reuniao.docx" --listar-figuras
python formatar_transcricao.py "reuniao.docx" --remover-ids 1 2 -o "reuniao_selecionada.docx"
```

Os IDs `1` e `2` são exemplos. Cada arquivo tem seus próprios IDs.
