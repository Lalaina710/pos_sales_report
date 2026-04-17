# Rapport Ventes PdV — Odoo 18

Export Excel des ventes Point de Vente avec wizard filtrable.

## Colonnes

| # | Colonne | Source |
|---|---------|--------|
| 1 | Date | `pos.order.date_order` (timezone user) |
| 2 | Code Articles | `product.product.default_code` |
| 3 | Code famille | `product.category` (dernier segment) |
| 4 | Famille | `product.category.name` |
| 5 | Designation Articles | `product.product.name` |
| 6 | PV | `pos.order.line.price_unit` |
| 7 | Point de vente | `pos.config.name` |
| 8 | Qtes | `pos.order.line.qty` |
| 9 | MTT HT | `pos.order.line.price_subtotal` |
| 10 | MTT TTC | `pos.order.line.price_subtotal_incl` |

## Fonctionnalites

- Header Famille en vert (style Sage)
- Autofiltre Excel sur toutes les colonnes
- Freeze panes (en-tete fige au scroll)
- Ligne TOTAL en pied (Qtes + HT + TTC)
- Negatifs en rouge (retours PdV)
- Timezone Madagascar correcte

## Filtres disponibles

- Periode (date debut / fin)
- Point(s) de vente
- Produits
- Familles produit

## Menu

**Point de Vente → Rapports → Analyse Ventes PdV**

## Installation

1. Copier dans `addons_path`
2. Mettre a jour la liste des modules
3. Installer « Rapport Ventes PdV (Analyse) »

## Dependances

- `point_of_sale`
- `xlsxwriter` (Python)

## Licence

LGPL-3

## Auteur

SOPROMER — [odoo.sopromer.mg](https://odoo.sopromer.mg)
