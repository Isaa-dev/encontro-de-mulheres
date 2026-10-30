# 1º Encontro Internacional de Mulheres – Inscrições

Site de inscrição gratuita para o encontro (06/11 às 19h e 07/11 às 17h, Rua Joaquim Afonso de Souza, 988), com divulgação do livro *Saindo das Sombras*, de Brenda Toet.

- Pede o nome completo, se a pessoa tem filhos e quantos filhos vão ao evento.
- Cada inscrição vai para uma planilha Google de controle.
- O limite é de **150 inscrições**: depois disso o site mostra "Inscrições encerradas".
- O mesmo nome não consegue se inscrever duas vezes.
- A coluna *Grupo do sorteio* separa as inscrições 1–50 (sorteio de sexta) e 51–150 (sorteio de sábado).

## Como funciona

| Parte | Onde fica |
|---|---|
| Página (`index.html`, `img/`) | GitHub Pages |
| Gravação na planilha (`apps-script/Code.gs`) | Google Apps Script, publicado como App da Web |

## Ligar o site à planilha

1. Em [script.google.com](https://script.google.com), crie um projeto e cole o conteúdo de `apps-script/Code.gs`.
2. **Implantar › Nova implantação › App da Web**. Em "Executar como", escolha **Eu**. Em "Quem pode acessar", escolha **Qualquer pessoa**. Depois autorize o acesso.
3. Copie o link gerado (termina em `/exec`).
4. No `index.html`, cole o link em `var API_URL = '';` e salve.

Enquanto `API_URL` estiver vazio, o site mostra "As inscrições abrem em breve" e não aceita envios.
