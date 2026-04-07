## 📊 Planilha de Exemplo: Controle de Metas e Vendas

Imagine que você é o analista responsável por acompanhar o desempenho da equipe comercial. Você tem a seguinte planilha (intervalo `A1:E6`):

| Vendedor | Vendas do Mês (R$) | Meta (R$) | % Atingido | Status |
| :--- | :--- | :--- | :--- | :--- |
| Ana Silva | R$ 55.000 | R$ 50.000 | 110% | Meta Batida |
| Bruno Costa | R$ 30.000 | R$ 40.000 | 75% | Atenção |
| Carlos Dias | R$ 15.000 | R$ 45.000 | 33% | Crítico |
| Diana Rocha | R$ 80.000 | R$ 60.000 | 133% | Meta Batida |
| Eduardo Luz | R$ 48.000 | R$ 50.000 | 96% | Atenção |

### 🎨 Aplicando a Formatação Condicional na Prática

Para transformar essa tabela estática em um **Dashboard Visual**, aplicaríamos as seguintes regras:

**1. Na Coluna "Vendas do Mês" (Escala de Cor):**
* Selecione de `R$ 55.000` até `R$ 48.000`.
* Aplique uma **Escala de Cor (Verde, Amarelo, Vermelho)**.
* *O que acontece:* O Excel pinta automaticamente os R$ 80.000 (maior valor) de verde forte, os R$ 15.000 (menor valor) de vermelho forte, e cria um degradê para os outros valores. Você bate o olho e vê quem vendeu mais.

**2. Na Coluna "% Atingido" (Conjunto de Ícones - Faróis):**
* Selecione de `110%` até `96%`.
* Aplique o **Conjunto de Ícones de Semáforo**.
* Configure as regras:
  * 🟢 Verde se for `>= 100%` (Bateu a meta).
  * 🟡 Amarelo se for `>= 80%` (Ficou perto).
  * 🔴 Vermelho se for `< 80%` (Ficou muito abaixo).
* *O que acontece:* Uma bolinha colorida aparece ao lado de cada porcentagem, mostrando visualmente a saúde da meta.

**3. Na Coluna "Status" (Texto que Contém):**
* Selecione de `Meta Batida` até `Atenção`.
* Aplique a regra **Texto que Contém...** "Meta Batida".
* Escolha a cor Verde.
* *O que acontece:* Apenas as células com a palavra exata "Meta Batida" ganharão um destaque verde, facilitando a visualização rápida para a gerência.

> 💡 **Lembre-se:** Se no mês seguinte o Carlos vender R$ 60.000, as cores, os faróis e os status vão mudar **sozinhos**. É isso que torna o Excel uma ferramenta indispensável no mercado!
