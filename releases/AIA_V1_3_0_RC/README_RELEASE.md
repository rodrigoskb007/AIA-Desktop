# AIA Desktop V1.3.0 RC

## Status da versão

Esta é a versão atualizada em validação (Release Candidate), publicada em 10 de novembro de 2026. A V1.2.0 permanece a última versão oficial estável; ela não é removida nem substituída por esta RC.

## Conteúdo

- `AIA_V1_3_0_RC.py`: código-fonte principal da nova versão.
- `INICIAR_AIA.bat`, `INSTALAR_DEPENDENCIAS_AIA.bat` e scripts de atalho: uso no Windows.
- `VALIDAR_AIA.py`, `VALIDAR_AIA.bat` e `SMOKE_TEST_UI.py`: validações reproduzíveis.
- `CHANGELOG_V1.3.0_RC.txt`: alterações desta RC.
- `RELATORIO_DE_VALIDACAO_V1.3.0_RC.txt`: resultado da validação local (31/31 verificações aprovadas).
- `Documentacao/`: guias e materiais de operação preservados da baseline V1.2.0.

## Evoluções principais

- D-I explicitamente destacado como indicador principal.
- PDF inicia com D-I e inclui a seção Dados Registrados.
- Threshold e confiança por classe de defeito, com snapshot por coleta.
- D-O passa a usar Correspondentes quando houver pareamento explícito; sem pareamento, permanece identificado como estimado.
- Exportações CSV/Excel com rastreabilidade por linha e auditoria.

## Limites da RC

Antes de classificar esta versão como oficial estável, executar a validação de campo prevista: importar planilha legada, exportar PDF/Excel/CSV e revisar visualmente o PDF resultante.

Arquivos de cache Python e dados de inspeção não são incluídos.
