# Changelog

Todas as alterações relevantes do LeadHunter são documentadas aqui.

## [2.0.0] - 2026-08-28 — Build comercial

### Adicionado
- Interface profissional com configurações de execução (limite de perfis por
  cidade, pausas configuráveis, opção de modo headless).
- Captura de e-mail e redes sociais em **lote** (até 3 requisições
  simultâneas) a partir do site público da empresa.
- Configurações persistentes em `config.json`.
- Relatórios Excel com 3 abas: *Leads Geral*, *Redes Sociais* e *Resumo*,
  com filtros automáticos, congelação do cabeçalho e largura de colunas.
- Logs rotativos em `logs/leadhunter.log`.
- Aviso de **uso responsável e LGPD** obrigatório antes de iniciar.
- Documentação: `README.md`, `TERMOS_DE_USO_E_PRIVACIDADE.md`, `CHANGELOG.md`.
- `requirements.txt` com dependências fixadas para builds reproduzíveis.
- `compilar.bat` para gerar o executável.

### Corrigido
- **Thread-safety da interface**: atualizações de widgets agora são
  encaminhadas para a thread principal via `self.after(...)`.
- **Slider de RAM enganoso removido**: a leitura era apenas informativa; agora
  a memória mostra o uso real do sistema e as opções impactam a execução.
- **Layout da interface**: corrigidos sobreposições de widgets/linhas.
- **Duplicação de enriquecimento**: o site não é mais consultado duas vezes
  pelo mesmo perfil.
- Seletores mais tolerantes à variação do layout do Google Maps, com
  detecção de possível bloqueio/CAPTCHA e mensagens objetivas.

### Melhorado
- Navegador com características mais discretas (user agent, locale, script de
  inicialização) e pauta consciente de uso responsável.
- Validação de e-mails (filtro de ruído como sentry, imagens, exemplos).
- Robustez na exportação e no tratamento de erros sem engolir falhas críticas.
