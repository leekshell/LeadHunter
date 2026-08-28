# LeadHunter — Prospecção Automática de Leads

> **v2.0 — build comercial.** Aplicação desktop para prospectar empresas por
> nicho e cidade, enriquecer os contatos com dados públicos do site da empresa
> e exportar os resultados em Excel com foco em WhatsApp e redes sociais.

![Estágio](https://img.shields.io/badge/est%C3%A1gio-Est%C3%A1vel-10b981)
![Licença](https://img.shields.io/badge/licen%C3%A7a-MIT-blue)

---

## 📌 Visão geral

O **LeadHunter** automatiza a etapa de prospecção B2B: o usuário informa um
nicho e uma ou mais cidades, e o software percorre as empresas encontradas,
captura nome, telefone, site, e-mail e redes sociais, remove duplicatas em um
banco SQLite local e gera relatórios Excel prontos para comercial.

A aplicação é **100% local**: os dados ficam no diretório de execução do
usuário e não são enviados para nenhum servidor.

## ✨ Funcionalidades

- 🔎 Busca por **nicho + múltiplas cidades** no Google Maps.
- 🖥️ **Interface desktop moderna** (CustomTkinter) com monitor em tempo real.
- 🎯 **Limite de perfis por cidade** e **pausas configuráveis** para uso
  responsável e redução de bloqueios.
- 📇 **Enriquecimento em lote** de e-mail, Instagram, Facebook e LinkedIn a
  partir do site público da empresa (limitado a 3 requisições simultâneas).
- 🧠 **Deduplicação** via SQLite (`INSERT OR IGNORE`) — leads já processados
  são pulados automaticamente.
- 📊 **Relatórios Excel** com 3 abas: *Leads Geral*, *Redes Sociais* e
  *Resumo* (com estatísticas e filtros nativos).
- 📝 **Logs rotativos** em `logs/leadhunter.log`.
- ⚙️ Configurações persistentes em `config.json`.
- 🧾 **Aviso de uso responsável e LGPD** integrado à interface.

## 🧰 Requisitos

- **Windows 10/11** (ambiente de produção recomendado). Também funciona em
  Linux/macOS com Python instalado.
- **Python 3.10+** (3.11 em diante é o recomendado).
- Conexão com a internet para a busca e enriquecimento.

## 🚀 Instalação

### Opção A — Rápida (Windows)

Execute o `instalador.bat` como administrador. Ele:

1. Cria um ambiente virtual Python (`.venv`).
2. Instala as dependências fixadas em `requirements.txt`.
3. Instala o Chromium do Playwright.
4. Abre o `App.py`.

### Opção B — Manual

```bash
python -m venv .venv
.venv\Scripts\activate            # Windows
source .venv/bin/activate         # Linux/macOS
pip install -r requirements.txt
playwright install chromium
python App.py
```

### Gerar executável (.exe)

Use o `compilar.bat` (Windows) ou rode:

```bash
pip install pyinstaller
pyinstaller --noconfirm --onefile --windowed --name LeadHunter App.py
```

O executável será criado em `dist/LeadHunter.exe`.

## 🖱️ Como usar

1. Informe o **nicho** (ex.: `Odontologia`).
2. Informe uma ou mais **cidades** separadas por vírgula (ex.:
   `Campinas, Sumaré, Artur Nogueira`).
3. Ajuste as **configurações de execução**:
   - `Máx. perfis/cidade`: limite por cidade.
   - `Pausa entre perfis (s)`: intervalo entre perfis (anti-bloqueio).
   - `Pausa entre cidades (s)`: intervalo entre cidades.
   - `Coletar dados do site`: habilita captura de e-mail.
   - `Coletar redes sociais`: habilita captura de Instagram/Facebook/LinkedIn.
4. **Confirme o uso responsável/LGPD**.
5. Clique em **Iniciar Varredura**.

Os relatórios são salvos em `relatorios/<data>/` e o histórico no banco
`leadhunter_storage.db`.

## ⚙️ Configuração

O arquivo `config.json` é criado no diretório da aplicação na primeira execução
e é editado pela interface. Campos principais:

| Chave                 | Descrição                                    | Padrão |
|-----------------------|----------------------------------------------|--------|
| `nicho`               | Nicho padrão                                 | `""`   |
| `cidades`             | Cidades padrão                               | `""`   |
| `max_perfis_por_cidade` | Limite de perfis por cidade                | `60`   |
| `delay_entre_perfis`  | Pausa mínima entre perfis (s)                | `1.2`  |
| `delay_entre_cidades` | Pausa entre cidades (s)                      | `8.0`  |
| `coletar_site`        | Captura e-mail do site                       | `true` |
| `coletar_redes`       | Captura redes sociais                        | `true` |
| `modo_headless`       | Navegador invisível                          | `true` |
| `confirmar_uso_responsavel` | Confirmação LGPD obrigatória           | `true` |

## 🏛️ Estrutura do projeto

```
LeadHunter/
├── App.py                         # Aplicação principal (interface + scraper)
├── instalador.bat                 # Instalação automática (Windows)
├── compilar.bat                   # Gera o executável (Windows)
├── requirements.txt               # Dependências fixadas
├── README.md                      # Documentação
├── TERMOS_DE_USO_E_PRIVACIDADE.md # Compliance / LGPD
├── LICENSE                        # MIT
├── logs/                          # Gerado em runtime
├── relatorios/                    # Gerado em runtime
├── config.json                    # Gerado em runtime
└── leadhunter_storage.db          # Gerado em runtime
```

## ⚠️ Uso responsável e conformidade

Este software deve ser usado exclusivamente para **contatos comerciais
legítimos**. A coleta automatizada de dados pode ser limitada pelos Termos de
Uso dos serviços consultados e, no Brasil, está sujeita à **LGPD**
(Lei 13.709/2018). O usuário é o controlador dos dados e deve:

- Usar apenas dados públicos com finalidade legítima e boa-fé.
- Manter registro das tratativas e conceder exclusão/retificação quando
  solicitado.
- Evitar spam, envio em massa não solicitado ou abordagens abusivas.
- Respeitar as taxas de acesso e as pausas configuradas.

Consulte [TERMOS_DE_USO_E_PRIVACIDADE.md](TERMOS_DE_USO_E_PRIVACIDADE.md)
antes de distribuir ou operar a ferramenta.

## 🗺️ Roadmap

- [ ] Perfil/contas de usuário e permissões.
- [ ] Filtros avançados (CNPJ, CEP, avaliação).
- [ ] Agendamento de varreduras recorrentes.
- [ ] Exportação para CRM (Pipedrive, HubSpot).
- [ ] Painel de métricas e dashboards.
- [ ] Automação de testes (Playwright + pytest).

## 📄 Licença

MIT — veja [LICENSE](LICENSE).

> **Aviso:** o software é fornecido "como está", sem garantias. A utilização é
> de inteira responsabilidade do usuário, que deve observar a legislação
> aplicável e os Termos de Uso das plataformas consultadas.
