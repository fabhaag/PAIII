# Automação de Certificados CMM - Schultz (MVP)

Este repositório contém a versão MVP (Minimum Viable Product) de uma aplicação web transacional simples desenvolvida para a **Schultz**. O objetivo do projeto é automatizar o processamento de relatórios de calibração de Máquinas de Medição por Coordenadas (CMM), consolidando o mapeamento de processos e os casos de uso estruturados para a disciplina de Projeto Aplicado III.

---

## 🎯 Sobre o Projeto

A aplicação resolve o problema do processamento manual de dados de calibração. Ela recebe planilhas operacionais em formato Excel (`.xlsx`) geradas pela máquina de medição (ex: CMM Contura da Carl Zeiss), processa os metadados e os resultados das medições, aplica as regras de negócio metrológicas e avalia automaticamente se o dispositivo foi **APROVADO** ou **REPROVADO**. Ao final, o sistema é capaz de gerar um certificado em PDF.

### 🏢 Contexto Acadêmico
* **Instituição:** Centro Universitário SENAI/SC - Campus Florianópolis
* **Disciplina:** Projeto Aplicado III (AV3)
* **Tutor:** Hugo Soares Barra
* **Data:** Outubro de 2026

### 👥 Equipe (Grupo 41)
* Nicolas Rafael Duma
* Fabiano Carcuchinski Haag
* Silas Bratti Batista
* Nicoly Pereira Carvalho
* Victor Hugo de Souza

---

## ⚙️ Funcionalidades e Regras de Negócio

1. **Ingestão e Tratamento de Dados:**
   * Leitura de planilhas Excel submetidas pelo usuário.
   * Extração de metadados do cabeçalho (Dispositivo de Controle, Data, Ordem, Metrologista, etc.).
   * Função auxiliar (`repair_xlsx_stream`) dedicada a corrigir problemas de *namespace* XML em arquivos exportados diretamente pelas máquinas de medição.
   * Limpeza e extração dos valores das características (Actual, Nominal, Upper Tol, Lower Tol).

2. **Cálculos Metrológicos Automáticos:**
   * **Média das Leituras:** Agrupamento dinâmico das $N$ medições realizadas e cálculo da média para cada ponto de ensaio.
   * **Desvio Padrão:** Cálculo estatístico baseado nas medições lidas.
   * **Erro Sistemático:** Cálculo da diferença absoluta entre a média lida e o valor padrão de referência (`abs(media_obtida - nominal)`).

3. **Validação e Lógica de Aprovação:**
   * **Validação por Ponto:** O sistema avalia individualmente cada ponto de ensaio, verificando se a média obtida está dentro dos limites admissíveis (Limite Inferior e Limite Superior).
   * **Feedback Visual:** Linhas da tabela recebem cores indicativas de status ("DENTRO DA TOLERÂNCIA" em verde, "FORA DA TOLERÂNCIA" em vermelho).
   * **Parecer Final:** Uma variável de controle monitora o processo. Se houver **qualquer** parâmetro fora da tolerância, o certificado inteiro recebe o status de **REPROVADO**. Caso contrário, é **APROVADO**.

4. **Geração de Certificado:**
   * Consolidação das informações e geração automática do certificado de calibração em formato PDF preenchido com as tabelas de resultados.

---

## 💻 Tecnologias Utilizadas

A aplicação foi desenvolvida em **Python** (IDE VS Code) com o uso intensivo das seguintes bibliotecas:
* **Pandas:** Para manipulação, limpeza, agrupamento e cálculos dos *dataframes* baseados no arquivo Excel.
* **Streamlit:** (Inferido pelas imagens/código) Para a construção ágil da interface web interativa (`st.session_state`, uploads, alertas visuais).
* **Openpyxl:** *Engine* para a leitura e manipulação das planilhas Excel.

---

## 🚀 Próximos Passos (Melhorias Futuras)

Embora os testes da versão MVP tenham sido satisfatórios, o projeto possui o seguinte *roadmap* de ajustes arquiteturais e de negócio previstos para as próximas iterações:

- [ ] **Ajuste de Incerteza:** Corrigir o cálculo da incerteza de medição ($U$) e do fator de abrangência ($k$), que na versão atual encontram-se com valores fixos (estáticos).
- [ ] **Refatoração de Arquitetura (MVC):** Melhorar a organização estrutural do código-fonte, quebrando o *script* monolítico atual e adotando o padrão **Model-View-Controller (MVC)** para separar responsabilidades lógicas e de interface.
- [ ] **Regras de Rejeição:** Refinar o entendimento e a lógica sistêmica sobre a emissão (ou não emissão) de certificados em PDF para ensaios classificados como reprovados.