# Relatório Analítico de Custo por Quilômetro — Grupo Maroni

Publica o dashboard de CPK da frota em um **link fixo** no GitHub Pages.
Toda vez que você atualizar as planilhas, o relatório é **regenerado e republicado sozinho** — o link nunca muda.

---

## O que tem neste pacote

| Arquivo | Para que serve |
|---|---|
| `dados/` | Pasta onde ficam **as duas planilhas do mês** (relatório de CPK + base de notas). |
| `gerar_relatorio.py` | Lê as planilhas e monta o `index.html`. Roda sozinho no GitHub. |
| `template.html` | O layout do dashboard (você não precisa mexer). |
| `.github/workflows/deploy.yml` | A automação que publica o relatório a cada atualização. |

---

## Configuração inicial (só uma vez)

1. Crie uma conta em **github.com** (se ainda não tiver).
2. Clique em **New repository**, dê o nome **`relatorio-cpk-maroni`**, deixe **Public** e crie.
3. Na página do repositório, clique em **Add file ▸ Upload files** e arraste **todo o conteúdo desta pasta** (incluindo a pasta `dados/` e a pasta oculta `.github/`). Confirme com **Commit changes**.
4. Vá em **Settings ▸ Pages** e, em **Source**, escolha **GitHub Actions**. Salve.
5. Abra a aba **Actions** e aguarde o fluxo "Publicar relatório de CPK" terminar (~2 minutos, fica verde ✅).

Pronto. Seu link fixo será:

```
https://SEU-USUARIO.github.io/relatorio-cpk-maroni/
```

(troque `SEU-USUARIO` pelo seu nome de usuário do GitHub). É esse link que você divulga para a equipe.

---

## Atualização mensal (o que você faz todo mês)

1. Prepare as **duas planilhas do mês** no seu computador:
   - o **Relatório de Custo por Quilômetro** (por placa, com os custos corrigidos);
   - a **Base de Dados de Custo de Manutenção** (as notas com datas).
2. No repositório do GitHub, entre na pasta **`dados/`**.
3. Clique em **Add file ▸ Upload files** e arraste as duas planilhas novas **por cima das antigas** (pode manter os mesmos nomes ou trocar — o sistema detecta as duas automaticamente). Se trocar o nome, apague as antigas.
4. Confirme com **Commit changes**.

Em ~2 minutos o relatório no **mesmo link** já mostra os números do mês novo. Não precisa mexer em mais nada.

> Dica: os nomes dos arquivos não importam. O gerador identifica sozinho qual é o relatório de CPK e qual é a base de notas pelas colunas de cada um.

---

## Rodar no seu computador (opcional, para testar antes)

Se quiser gerar o `index.html` localmente antes de subir:

```bash
pip install pandas openpyxl
python gerar_relatorio.py
```

Abra o `index.html` gerado no navegador.

---

## Agrupar novos modelos (opcional)

Se aparecerem modelos escritos de formas diferentes que devem contar como o mesmo,
edite o trecho `OVERRIDE` no topo do `gerar_relatorio.py`. Exemplo já configurado:

```python
OVERRIDE = {
    'SCANIA RH 450 PLUS': 'SCANIA R450 PLUS',
    'SCANIA RN 450 PLUS': 'SCANIA R450 PLUS',
    'SCANIA RH 500 NA':   'SCANIA R500 NA',
    'SCANIA R500 NA':     'SCANIA R500 NA',
    'MBATEGO 2429 - SIDER': 'MB ATEGO 2429',
}
```

À esquerda vai o nome **exato** como está na planilha (em MAIÚSCULAS); à direita, o nome agrupado.
