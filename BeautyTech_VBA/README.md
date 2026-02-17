# 💅 BeautyTech VBA
**Desenvolvido por LogicUp Solutions**

O **BeautyTech** é um sistema de gestão completo e independente desenvolvido em Excel/VBA, projetado especificamente para Salões de Beleza e Barbearias. 
Focado em alta performance e UX/UI, o projeto rompe as barreiras do VBA tradicional, entregando uma interface moderna, responsiva e com experiência de uso semelhante a aplicações web (SaaS).

## 🏗️ Módulos do Sistema

### 🔐 01. Controle de Acesso & Segurança
- **Instância Isolada:** O sistema roda em uma janela independente, ocultando a interface padrão do Excel.
- **Níveis de Acesso:** Permissões baseadas em hierarquia (Administrador com visão financeira vs. Profissional com foco na agenda).
- **Log de Auditoria:** Rastreamento estruturado de acessos e operações críticas.

### 📅 02. Agenda Visual Inteligente
- **Dashboard Calendarizado:** Visualização de horários limpa e dinâmica, sem dependência de controles ActiveX instáveis.
- **Status Colorizados:** Gestão visual instantânea de agendamentos (Confirmado, Pendente, Cancelado).
- **Bloqueio Automático:** Motor de inteligência de tempo para prevenir sobreposição de horários.

### 💰 03. Financeiro & Comissionamento
- **Fluxo de Caixa Automático:** Lançamento de receitas integrado diretamente à finalização dos atendimentos.
- **Divisão de Lucros (Split):** Cálculo automatizado de comissões por profissional e por serviço.
- **DRE Simplificado:** Geração de gráficos e indicadores de Receita x Despesa x Lucro Líquido.

### 📦 04. Estoque Técnico (Diferencial)
- **Ficha Técnica de Serviços:** Baixa automática e fracionada de insumos (ex: gramas de gel, pares de luvas) ao concluir um atendimento.
- **Alertas Inteligentes:** Notificações automáticas de estoque em nível crítico logo na inicialização do sistema.

### 📱 05. Integração WhatsApp Omnichannel
- **Lembretes Automáticos:** Disparo de confirmações de agendamento via API oficial do WhatsApp Web.
- **Redução de No-Show:** Fluxo de comunicação automatizado e direto com o cliente para maximizar a ocupação da agenda.

## 🗺️ Roadmap de Desenvolvimento (4 Semanas)

- [x] **Semana 1: Arquitetura, UX/UI e Conexão de Dados**
  - [x] Construção do Design System SaaS (Telas de Login e Dashboard).
  - [x] Menu lateral responsivo com navegação fluida e tipografia avançada.
  - [x] Estruturação da conexão segura com o Banco de Dados (Access/SQL).

- [ ] **Semana 2: O *Core* (Agenda Inteligente e Cadastros)**
  - Lógica de roteamento dinâmico (MultiPage).
  - CRUD de Clientes, Profissionais e Serviços (usando ListView customizado).
  - Motor visual da Agenda com bloqueio automático de conflitos de horários.

- [ ] **Semana 3: Inteligência de Estoque e Motor Financeiro**
  - Engenharia da Ficha Técnica (baixa fracionada e automática de insumos).
  - Módulo de Caixa (Entradas e Saídas integradas aos atendimentos).
  - Cálculo de *Split* de Comissões e geração do DRE dinâmico.

- [ ] **Semana 4: Automação e Empacotamento Comercial**
  - Integração com API do WhatsApp Web (Lembretes automáticos e antino-show).
  - Refinamento dos logs de auditoria e segurança de acessos.
  - Testes finais, ofuscação de código e empacotamento para comercialização.
  
  <br>
<div align="center">
  <b>Desenvolvido com 💻 e ☕ por Danyllo Pereira</b><br>
  <i>Especialista VBA e Fundador da <b>LogicUp Solutions</b> 🚀</i><br><br>
</div>