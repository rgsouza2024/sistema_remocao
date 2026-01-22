
try:
    from motor_remocao import processar_remocao, normalizar_colunas, aplicar_congelamento
    import app
    print("✅ Sucesso: motor_remocao e app importados corretamente.")
except ImportError as e:
    print(f"❌ Erro de Importação: {e}")
except Exception as e:
    print(f"❌ Erro Genérico: {e}")
