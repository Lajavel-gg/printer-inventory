# 🖨️ Printer Inventory

Script PowerShell pour inventorier automatiquement les imprimantes réseau (Zebra et Printer) sur un domaine Active Directory.

## 📋 Description

Ce script interroge tous les postes d'une OU Active Directory via WinRM pour récupérer les imprimantes réseau installées dans les profils utilisateurs. Il génère un fichier Excel propre avec **une ligne par imprimante** et la liste de tous les postes qui l'utilisent.

À chaque exécution, il complète le fichier existant avec les nouvelles données trouvées — idéal pour scanner progressivement un parc avec des postes pas toujours allumés.

## ✅ Prérequis

- PowerShell 5.1 ou supérieur
- Module Active Directory (`RSAT-AD-PowerShell`)
- WinRM activé sur les postes clients
- Droits administrateur sur les postes du domaine
- Microsoft Excel installé sur le poste depuis lequel le script est lancé

## ⚙️ Configuration

Avant de lancer le script, modifiez les deux variables en haut du fichier :

```powershell
$SearchBase = "OU=TON_OU,OU=Worldwide_Sites,OU=TON_DOMAINE,DC=TON_DOMAINE,DC=pri"
$xlsxPath   = "$env:USERPROFILE\Desktop\Inventaire_Imprimantes_Final.xlsx"
```

- `$SearchBase` : l'OU Active Directory contenant les postes à scanner
- `$xlsxPath` : chemin du fichier Excel généré (par défaut sur le Bureau)

## 🚀 Utilisation

1. Ouvrir PowerShell en tant qu'administrateur
2. Modifier `$SearchBase` avec votre OU
3. Lancer le script :

```powershell
.\Inventaire_Imprimantes.ps1
```

Le script affiche en temps réel les nouvelles imprimantes trouvées en vert et génère le fichier Excel à la fin.

## 📁 Format de sortie

| Nom imprimante | Postes |
|----------------|--------|
| zebra-01 | DESKTOP-001, LAPTOP-002, LAPTOP-003 |
| printer-07 | DESKTOP-010, LAPTOP-011 |

## 🔄 Fonctionnement

1. **Chargement** — Si un fichier Excel existe déjà, il charge les données existantes
2. **Scan** — Interroge tous les postes de l'OU via WinRM en parallèle (50 postes simultanés)
3. **Nettoyage** — Supprime les doublons, les redirections RDP et les imprimantes personnelles
4. **Export** — Génère ou met à jour le fichier Excel avec une ligne par imprimante

## 💡 Conseils

- Relancez le script plusieurs fois sur plusieurs jours pour maximiser la couverture (tous les postes ne sont pas allumés en même temps)
- Lancez-le de préférence en heure de pointe (9h-10h) pour avoir le maximum de postes allumés
- Le script ne supprime jamais les données existantes, il complète uniquement

## 📝 License

MIT
