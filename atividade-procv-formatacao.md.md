# 💻 Laboratório Prático: Automação de Estoque na "TechMinds"

**Disciplina:** PTIC
**Técnicas Abordadas:** `PROCV` (Busca de Dados em Matrizes) e Formatação Condicional (Front-End / Alertas Visuais).

## 🏢 O Cenário (Seu Desafio no Mercado de Trabalho)

Você acabou de ser contratado como Analista de Dados da **TechMinds**, uma loja de hardwares e equipamentos de TI. 

O gerente de suprimentos tem um grande problema: o controle de estoque diário é feito de forma manual. Quando um cliente pergunta o preço de um produto, o vendedor perde muito tempo procurando na lista. Ele te entregou a **Base de Dados** oficial da loja e pediu para você criar um **Painel Automatizado**. 

A regra de negócio é clara: quando o operador digitar o Código do Produto, o sistema (Excel) deve varrer o banco de dados e puxar o Nome e o Custo automaticamente. Além disso, o painel deve alertar visualmente (com cores) se o estoque estiver perigosamente baixo.

---

## 🛠️ Passo 1: Montando o Banco de Dados (O "Back-End")

Abra uma nova planilha no Excel. Na "Planilha 1", vamos criar o nosso Banco de Dados oficial. Preste muita atenção na digitação dos códigos (nossa Chave Primária).

**Crie esta tabela no intervalo de `A1` até `C41`:**

| Código | Produto | Custo Unitário (R$) |
| :--- | :--- | :--- |
| **HW-001** | SSD 1TB NVMe | 350,00 |
| **HW-002** | Memória RAM 16GB DDR4 | 280,00 |
| **HW-003** | Processador Core i5 12ª Geração | 950,00 |
| **HW-004** | Processador Ryzen 5 5600X | 890,00 |
| **HW-005** | Placa-Mãe B550M | 720,00 |
| **HW-006** | Placa de Vídeo RTX 3060 12GB | 1.850,00 |
| **HW-007** | Placa de Vídeo RX 6600 8GB | 1.450,00 |
| **HW-008** | Fonte de Alimentação 650W 80 Plus | 340,00 |
| **HW-009** | Gabinete ATX com Vidro Temperado | 260,00 |
| **HW-010** | Cooler para Processador RGB | 130,00 |
| **HW-011** | Water Cooler 240mm | 380,00 |
| **HW-012** | HD Interno 2TB SATA III | 310,00 |
| **AC-010** | Teclado Mecânico Switch Azul | 180,00 |
| **AC-011** | Mouse Óptico Sem Fio | 90,00 |
| **AC-012** | Headset Gamer 7.1 Surround | 220,00 |
| **AC-013** | Webcam Full HD 1080p | 160,00 |
| **AC-014** | Mousepad Gigante (90x40cm) | 65,00 |
| **AC-015** | Suporte Articulado para Monitor | 195,00 |
| **AC-016** | Cadeira Gamer Ergonômica | 950,00 |
| **AC-017** | Microfone Condensador USB | 290,00 |
| **AC-018** | Caixa de Som Bluetooth 2.1 | 140,00 |
| **AC-019** | Teclado Padrão ABNT2 Escritório | 45,00 |
| **AC-020** | Pen Drive 64GB USB 3.0 | 55,00 |
| **NT-100** | Roteador Wi-Fi 6 Gigabit | 450,00 |
| **NT-101** | Switch de Rede 8 Portas | 100,00 |
| **NT-102** | Cabo de Rede RJ45 Cat6 (10 metros) | 48,00 |
| **NT-103** | Adaptador Wi-Fi USB Dual Band | 60,00 |
| **NT-104** | Repetidor de Sinal Wi-Fi | 115,00 |
| **SW-001** | Licença Sistema Operacional Win 11 Pro | 850,00 |
| **SW-002** | Licença Pacote Office 365 (Anual) | 360,00 |
| **SW-003** | Licença Antivírus Premium (Anual) | 85,00 |
| **MO-001** | Monitor LED 24" Full HD | 720,00 |
| **MO-002** | Monitor Gamer 27" 144Hz | 1.250,00 |
| **MO-003** | Monitor Ultrawide 29" | 1.150,00 |
| **PR-001** | Impressora Multifuncional Wi-Fi | 680,00 |
| **PR-002** | Cartucho de Tinta Preto Original | 90,00 |
| **PR-003** | Cartucho de Tinta Colorido Original | 105,00 |
| **CB-001** | Cabo HDMI 2.0 (2 metros) | 30,00 |
| **CB-002** | Adaptador USB-C para HDMI | 70,00 |
| **EN-001** | Nobreak 1200VA | 620,00 |

---

## 🖥️ Passo 2: O Painel do Operador (O "Front-End")

Crie uma "Planilha 2" (ou pule algumas colunas) e monte a tabela onde o usuário vai trabalhar de fato. Deixe as linhas abaixo dos títulos em branco, pois é lá que as fórmulas vão agir.

**Crie esta tabela no intervalo `E1:H6`:**

| Cód. Digitado | Produto Retornado | Estoque Atual | Status do Estoque (Cor) |
| :--- | :--- | :--- | :--- |
| HW-006 | *(Sua Fórmula PROCV)* | 15 | *(Formatação aqui)* |
| NT-104 | *(Sua Fórmula PROCV)* | 4 | *(Formatação aqui)* |
| SW-002 | *(Sua Fórmula PROCV)* | 8 | *(Formatação aqui)* |
| AC-016 | *(Sua Fórmula PROCV)* | 50 | *(Formatação aqui)* |
| PR-003 | *(Sua Fórmula PROCV)* | 9 | *(Formatação aqui)* |

---

## ⚙️ Passo 3: Missão PROCV 

**Sua tarefa:** Na coluna **Produto Retornado**, crie uma fórmula `PROCV` para que o Excel leia o "Cód. Digitado" na mesma linha, varra o Banco de Dados (Passo 1), encontre a descrição correta e a retorne na tela.

> 💡 **Dica de Lógica:** Como nossa matriz de busca tem 40 linhas, é vital travar as referências (`$`)! Se você não travar a matriz (ex: `$A$1:$C$41`) e arrastar a fórmula para baixo, a matriz vai escorregar (virando `A2:C42`, `A3:C43`...) e produtos começarão a dar erro `#N/D` (Não Disponível).

---

## 🚦 Passo 4: Missão Formatação Condicional (Lógica `SE` Visual)

O gerente definiu a seguinte regra de negócio:
> *"Qualquer produto que tenha um estoque **menor que 10 unidades** está em nível crítico e precisamos comprar urgentemente."*

**Sua tarefa:**
1. Selecione os números da coluna **Estoque Atual** (os números que estão na tabela do Passo 2).
2. Na guia superior do Excel, encontre a ferramenta **Formatação Condicional**.
3. Crie uma regra para que o preenchimento da célula fique **Vermelho** apenas se o número for **menor que 10**.

🔥 **Desafio Extra:** Altere manualmente o estoque do produto `AC-016` de 50 para 5. A célula deve ficar vermelha automaticamente. Se isso acontecer, seu sistema está validado e pronto para o mercado!

---


