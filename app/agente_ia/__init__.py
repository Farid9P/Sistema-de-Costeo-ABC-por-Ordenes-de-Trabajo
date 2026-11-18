"""Agente IA de Finanzas y Contabilidad: interpreta preguntas, responde leyendo
la base SQLite y propone mejoras. Es de solo lectura."""

from .agente import AVISO_SOLO_LECTURA, AgenteFinanzas, ResultadoAgente

__all__ = ["AgenteFinanzas", "ResultadoAgente", "AVISO_SOLO_LECTURA"]
