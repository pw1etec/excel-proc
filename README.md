# 🔍 Funções de Busca no Excel (PROCV, PROCH e PROCX) — Guia Definitivo e Didático

**Por: Professor Alessandro**
*(Material de Apoio — PTIC)*

Se você quer trabalhar com dados de forma profissional, dominar as funções da família **PROC** (Procurar) é obrigatório. Elas são a espinha dorsal de relatórios no mercado de trabalho — de finanças e RH a logística e Business Intelligence (BI).

💡 **A Visão do Programador:** No fundo, um PROCV nada mais é do que um **algoritmo de busca linear**. Você entrega uma **Chave Primária** (um ID, CPF ou código único), a função percorre a matriz de dados (array) até encontrar a correspondência exata e retorna o valor atrelado àquele registro. É o conceito de Banco de Dados Relacional rodando direto na sua planilha!

---

## 🔹 1. PROCV (Procura Vertical)

### 📍 O Conceito
O **PROCV** busca um valor na **primeira coluna** de uma tabela (de cima para baixo) e retorna uma informação que está na mesma linha, em outra coluna.

### 🧾 A Sintaxe (A Anatomia da Função)
```excel
=PROCV(valor_procurado; matriz_tabela; num_indice_coluna; [procurar_intervalo])
```

**Traduzindo os parâmetros:**
* `valor_procurado`: O que você quer encontrar? (Sua chave de busca).
* `matriz_tabela`: Onde está o banco de dados? *(Regra: a 1ª coluna deve conter a chave).*
* `num_indice_coluna`: Qual coluna tem a resposta que você quer? (Contando a partir de 1).
* `[procurar_intervalo]`: A busca é exata? Use `0` ou `FALSO` (99% dos casos no mercado).

### 💼 Exemplo Prático (Sistema de Vendas)
Temos a nossa base de dados de produtos:

| Código (Col 1) | Produto (Col 2) | Preço (Col 3) |
| :--- | :--- | :--- |
| 101 | Teclado Mecânico | R$ 150,00 |
| 102 | Mouse Sem Fio | R$ 80,00 |
| 103 | Monitor 24" | R$ 800,00 |

Você quer buscar o preço do código **102**.
```excel
=PROCV(102; A2:C4; 3; FALSO)
```
👉 **Resultado:** `80,00`

> 🧠 **Tradução Mental do Algoritmo:**
> "Excel, pegue o código 102, varra a primeira coluna. Quando achar, ande até a coluna 3 e me devolva o valor exato."

### ⚠️ Erros Comuns no PROCV
* ❌ A chave de busca **não está na primeira coluna** da matriz selecionada.
* ❌ Esquecer de usar o `FALSO` (ou `0`) no final, trazendo resultados aproximados incorretos.
* ❌ Esquecer de travar a matriz com cifrão (`$`) antes de arrastar a fórmula.

---

## 🔹 2. PROCH (Procura Horizontal)

### 📍 O Conceito
A lógica é idêntica ao PROCV, mas a varredura muda de eixo. Ele busca na **primeira linha** (da esquerda para a direita) e retorna valores das linhas abaixo.

### 🧾 A Sintaxe
```excel
=PROCH(valor_procurado; matriz_tabela; num_indice_linha; [procurar_intervalo])
```

### 💼 Exemplo Prático (RH / Metas Mensais)

| | Jan (B) | Fev (C) | Mar (D) |
| :--- | :--- | :--- | :--- |
| **Meta (Linha 1)** | 100 | 120 | 150 |
| **Real (Linha 2)** | 90 | 110 | 140 |

Você quer buscar a meta batida (Real) do mês de **Fev**.
```excel
=PROCH("Fev"; B1:D2; 2; FALSO)
```
👉 **Resultado:** `110`

---

## 🏆 Resumo e Comparação

| Função | Eixo de Busca | Complexidade | Uso no Mercado |
| :--- | :--- | :--- | :--- |
| **PROCV** | Vertical (↓) | Média | Padrão absoluto. Obrigatório saber. |
| **PROCH** | Horizontal (→) | Média | Uso situacional (tabelas deitadas). |
| **PROCX** | Qualquer Direção (⬌) | Baixa / Ágil | O futuro. Mais rápido e seguro. |

---

## 💡 Boas Práticas (Nível Profissional)

1. **🔒 Garanta a Busca Exata:** Acostume-se a sempre usar `0` ou `FALSO` no último argumento do PROCV/PROCH.
2. **📌 Trave as Matrizes (Referências Absolutas):** Ao selecionar o banco de dados, aperte `F4` para inserir os cifrões (ex: `$A$1:$C$100`). Isso congela a matriz e impede que ela "escorregue" quando você arrastar a fórmula para outras linhas.
3. **🧹 Limpe seus Dados:** O Excel entende "101" (Texto) e 101 (Número) como coisas diferentes. Se o PROCV der erro `#N/D`, verifique se as chaves primárias estão no mesmo formato de dados.
