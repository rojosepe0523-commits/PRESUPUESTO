# Apps Script de la hoja de presupuesto

`Code.gs` es la API que usa la app (versión 4). Reemplaza al `Code.gs` anterior.

## Instalar o actualizar (una sola vez, unos 3 minutos)

1. En la hoja: **Extensiones > Apps Script**.
2. Abre `Code.gs`, borra todo su contenido y pega el de este repositorio.
3. (Opcional, para la recomendación con IA) **Configuración del proyecto** (engranaje) > **Propiedades del script** >
   Agregar propiedad: `ANTHROPIC_API_KEY` = tu clave de https://console.anthropic.com. Sin clave, la recomendación se arma con reglas fijas.
4. Arriba elige la función **configurar** y pulsa **Ejecutar**. Acepta los permisos (hoja, activadores y conexión externa para la IA).
5. **Implementar > Gestionar implementaciones** > lápiz > Versión: **Nueva versión** > **Implementar**. La URL `/exec` no cambia.
6. Abre la app; si dice "Clave incorrecta", ejecuta **mostrarClave** y pega la clave en Ajustes.

## Qué hace

- **Catálogo dinámico**: crear, editar, renombrar (también en los gastos ya registrados) y borrar categorías y métodos de pago desde la pestaña Categorías de la app. Cada categoría tiene un tipo: fija, esencial, opcional o ahorro.
- **Reinicio por salario**: al registrar un salario empieza un nuevo mes en $0 y el anterior se guarda en la pestaña **Cierres** (ingresos, gastos, saldo, tasa de ahorro y gasto por categoría, con gráfico de evolución).
- **Recomendación mensual**: un minuto después de cerrar el mes se genera la recomendación (con Claude si hay clave). Un activador diario a las 6 a. m. la genera si faltó. Se ve en el Dashboard de la app y en Cierres.

## Propiedades opcionales

| Propiedad | Para qué | Valor por defecto |
|---|---|---|
| `ANTHROPIC_API_KEY` | Activa la recomendación con IA | (vacía: reglas fijas) |
| `IA_MODELO` | Modelo de Claude | `claude-opus-5-5` |
| `IA_ESFUERZO` | `low`, `medium` o `high` | `medium` |
| `IA_ENVIAR_DESCRIPCIONES` | `no` para no enviar las descripciones de tus gastos a la IA | se envían |

No ejecutes `RepararHoja` de nuevo: reescribiría el catálogo.
