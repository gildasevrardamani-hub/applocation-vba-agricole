# Matrice des arbitrages métier — AGRIECO PRO

## 1. Objet du document

Ce document recense les paramètres métier, choix de calcul et décisions fonctionnelles à valider avant le développement d'AGRIECO PRO.

Il complète les documents suivants :

- `ARCHITECTURE_AGRIECO_PRO.md` ;
- `DECISIONS_PRODUIT_V1_AGRIECO_PRO.md` ;
- `REGLES_CALCUL_AGRIECO_PRO.md`.

Il ne contient aucun code VBA, ne crée aucun UserForm et ne modifie pas l'architecture physique du futur classeur Excel.

## 2. Statuts de validation

| Statut | Signification |
|---|---|
| À valider | Décision nécessaire avant développement. |
| À préciser | Principe accepté, mais paramètres ou seuils à détailler. |
| Recommandé | Proposition recommandée par Codex, non encore validée. |
| Validé | Décision officiellement approuvée. |
| Reporté | Sujet préparé mais non obligatoire pour la V1. |

## 3. Niveaux de priorité

| Priorité | Signification |
|---|---|
| Critique | Bloque ou sécurise directement le moteur économique V1. |
| Haute | Fort impact sur les calculs, rapports ou décisions utilisateur. |
| Moyenne | Peut être ajusté pendant la V1 sans remettre en cause le socle. |
| Basse | Peut être reporté sans risque majeur pour la V1. |

---

# 4. Matrice synthétique des arbitrages

| N° | Thème | Règle concernée | Proposition actuelle | Options possibles | Avantages et limites des options | Recommandation de Codex | Impact sur la V1 | Impact sur versions futures | Priorité | Statut |
|---:|---|---|---|---|---|---|---|---|---|---|
| 1 | Seuils d'alerte | Déclenchement des alertes économiques | Seuils paramétrables par indicateur | A. Seuils standards uniques ; B. Seuils par spéculation ; C. Seuils par zone et spéculation | A simple mais peu précis ; B meilleur métier mais plus de paramétrage ; C très précis mais lourd | Retenir B pour la V1, préparer C | Nécessite `T_Param_Alertes` simple | Extension multi-zone facilitée | Critique | À valider |
| 2 | Résultat négatif | Alerte déficit | Alerte si résultat financier ou économique < 0 | A. Sur résultat économique seulement ; B. Sur résultat financier seulement ; C. Sur les deux | A mesure rentabilité complète ; B mesure cash réel ; C plus complet mais plus d'alertes | Retenir C avec libellés distincts | Rapports plus explicites | Compatible finance et conseil | Critique | Recommandé |
| 3 | Marge faible | Alerte de marge insuffisante | Seuil paramétrable | A. Marge nette/produit brut < 10 % ; B. seuil par spéculation ; C. seuil défini par organisation | A simple ; B pertinent techniquement ; C flexible mais moins comparable | B avec valeur par défaut 10 % si absent | Demande paramètre par spéculation | Ajustable par filière | Haute | À valider |
| 4 | Coût anormal | Coût unitaire ou coût/ha incohérent | Comparaison à un coût de référence | A. Aucune alerte ; B. seuil ±20 % ; C. seuil paramétrable | A simple mais faible contrôle ; B opérationnel ; C robuste | C, avec valeurs initiales par défaut | Améliore diagnostic | Prépare référentiels régionaux | Haute | À préciser |
| 5 | Rendement incohérent | Rendement trop bas ou trop élevé | Comparaison à rendement de référence | A. seuil unique ; B. seuil par spéculation ; C. seuil par zone/campagne | A simple ; B adapté V1 ; C plus précis mais plus lourd | B pour V1 | Contrôle qualité immédiat | Extension agronomique future | Haute | À valider |
| 6 | Prix incohérent | Prix très différent du prix de référence | Seuil paramétrable | A. alerte à ±30 % ; B. seuil par produit ; C. seuil par marché/date | A simple ; B bon compromis ; C précis mais nécessite données | B pour V1 | Détecte erreurs de saisie | Prépare veille de marché | Haute | À valider |
| 7 | Besoin de trésorerie élevé | Besoin maximal / produit brut ou investissement | Seuil paramétrable | A. montant absolu ; B. ratio sur produit brut ; C. ratio sur investissement ; D. combinaison | A peu comparable ; B lisible ; C utile financement ; D complet mais complexe | D dans le rapport, B comme alerte V1 | Aide décision de financement | Prépare scoring finance | Haute | À préciser |
| 8 | Pondérations IPE | Score final IPE sur 100 | Pondérations provisoires paramétrables | A. conserver pondérations proposées ; B. pondérations par profil ; C. pondérations par filière | A rapide ; B personnalisé ; C précis mais lourd | A pour démarrer, historiser, ouvrir B plus tard | Classement opérationnel | Personnalisation future | Critique | À valider |
| 9 | Critères IPE | Rentabilité, durée, investissement, risque, trésorerie, marché, main-d'œuvre | 7 critères retenus | A. 7 critères ; B. réduire à 5 ; C. ajouter critères sociaux/environnementaux | A équilibré ; B plus simple ; C plus complet mais hors V1 | A pour V1 | Cohérent avec règles actuelles | Extensions possibles | Critique | Recommandé |
| 10 | Normalisation IPE | Transformation 0-100 | Min/max paramétrables | A. bornes fixes ; B. bornes observées ; C. bornes référentielles | A simple ; B relatif au groupe ; C plus stable | C si référentiel disponible, sinon A | Nécessite paramétrage | Comparaisons historiques fiables | Critique | À préciser |
| 11 | Classe IPE | Classes A à E | A 80-100, B 65-79, C 50-64, D 35-49, E 0-34 | A. garder seuils ; B. seuils paramétrables ; C. seuils par profil | A simple ; B flexible ; C complexe | B avec valeurs proposées par défaut | Rapports clairs | Adaptable par institution | Haute | À valider |
| 12 | Historique IPE | Version des pondérations | Historiser les pondérations utilisées | A. pas d'historique ; B. historique simple ; C. historique détaillé avec auteur/date/commentaire | A non traçable ; B suffisant ; C robuste mais plus lourd | B en V1, préparer C | Garantit reproductibilité | Audit avancé futur | Critique | Recommandé |
| 13 | Pondérations IQD | Complétude 40, cohérence 30, précision 20, traçabilité 10 | Pondération proposée | A. conserver ; B. modifier ; C. pondération par type de compte | A simple ; B ajustable ; C précis mais lourd | A pour revue, validation métier requise | Score qualité immédiat | Ajustable par contexte | Haute | À valider |
| 14 | Pénalité saisie simplifiée | IQD ajusté selon niveau de détail | Pénalité de principe prévue | A. aucune pénalité ; B. pénalité fixe ; C. pénalité selon champs estimés | A surévalue qualité ; B simple ; C plus juste | C si possible, B au démarrage | Différencie comptes simples/détaillés | Améliore fiabilité | Haute | À préciser |
| 15 | Seuil minimal IQD | Affichage des recommandations | À définir | A. pas de seuil ; B. avertissement <70 % ; C. blocage recommandation forte <70 % | A risqué ; B souple ; C sécurisant | C pour recommandations fortes, B pour rapports | Évite décisions sur données faibles | Utile diagnostic automatique | Haute | À valider |
| 16 | Main-d'œuvre familiale | Valorisation économique | Valorisation possible, exclue du financier si non payée | A. non valorisée ; B. valorisée au barème unique ; C. barème par tâche/zone | A sous-estime coûts ; B simple ; C plus précis mais lourd | B pour V1, préparer C | Résultat économique crédible | Analyse sociale/emploi future | Critique | À valider |
| 17 | Barème main-d'œuvre | Coût journalier de référence | À définir | A. barème unique FCFA/jour ; B. par opération ; C. par zone/saison/profil | A rapide ; B pertinent ; C très fin mais complexe | A pour V1 avec option observation | Simplifie saisie | Paramétrage avancé possible | Critique | À préciser |
| 18 | Main-d'œuvre salariée | Charge réellement payée | Charge payée si salaire versé | A. montant global ; B. jours × coût/jour ; C. détail par opération | A rapide ; B calculable ; C analytique | B en V1 | Alimente coûts et trésorerie | Analyse opérations future | Haute | Recommandé |
| 19 | Autoconsommation | Produit économique séparé des ventes | Valoriser et intégrer au produit brut économique | A. exclure ; B. inclure au prix moyen de vente ; C. inclure au prix marché local | A sous-estime valeur ; B simple ; C plus exact | C si prix marché connu, sinon B | Produit brut économique fiable | Compatible sécurité alimentaire | Critique | À valider |
| 20 | Dons et usages familiaux | Flux non vendus | À classer | A. autoconsommation ; B. produit secondaire ; C. sortie non valorisée | A souvent pertinent ; B rare ; C sous-estime | A avec sous-type `Don/consommation` | Clarifie flux physiques | Reporting social futur | Moyenne | À préciser |
| 21 | Prix de référence | Prix par spéculation, zone, campagne | Référentiel paramétrable | A. prix unique ; B. prix par spéculation ; C. prix par zone/campagne/marché | A simple ; B bon V1 ; C plus précis mais exige données | B pour V1, préparer C | Sert alertes prix/autoconsommation | Veille de marché future | Haute | À valider |
| 22 | Rendement de référence | Rendement attendu ou standard | Référentiel paramétrable | A. rendement unique ; B. par spéculation ; C. par zone/système | A peu fiable ; B acceptable ; C meilleur mais plus lourd | B pour V1 | Prévisions simples | Extension agronomique | Haute | À valider |
| 23 | Sources des références | Origine prix/rendements/barèmes | Source à documenter | A. source libre ; B. source obligatoire ; C. source + date + validateur | A simple ; B traçable ; C robuste | B pour V1, C pour paramètres critiques | Améliore confiance | Audit futur | Moyenne | À préciser |
| 24 | Durées d'amortissement | Durée par actif | Durée paramétrable | A. durée saisie libre ; B. référentiel par catégorie ; C. référentiel verrouillé | A flexible mais risqué ; B équilibré ; C fiable mais rigide | B | Calcul amortissement stable | Harmonisation multi-clients | Critique | À valider |
| 25 | Valeur résiduelle | Amortissement linéaire | À confirmer | A. toujours zéro ; B. saisie facultative ; C. référentiel par actif | A simple ; B flexible ; C plus précis | B avec zéro par défaut | Simple et adaptable | Gestion actifs future | Moyenne | À valider |
| 26 | Prorata amortissement | Cycles inférieurs à un an | Appliquer fraction de période | A. annualiser seulement ; B. prorata mois ; C. prorata jours | A simple mais imprécis ; B bon compromis ; C précis mais plus de dates | B pour V1 | Cohérent cycles | Plus précis futur | Haute | À valider |
| 27 | Unités locales | Gestion unités et conversions | Référentiel d'unités | A. unités standards seulement ; B. unités locales avec coefficients ; C. coefficients par zone/marché | A limite terrain ; B utile ; C précis mais lourd | B | Saisie réaliste | Extension régionale | Critique | À valider |
| 28 | Coefficients conversion | Conversion vers unité référence | Coefficients validés | A. coefficient global ; B. coefficient par produit ; C. coefficient par produit/zone | A risqué ; B bon V1 ; C plus précis | B | Évite erreurs sacs/paniers | Adaptable marché local | Critique | À préciser |
| 29 | Stock d'intrants | Gestion mode individuel | Fonction demandée | A. suivi simple quantité/valeur ; B. lots et dates ; C. inventaire complet | A léger ; B utile mais plus lourd ; C hors V1 | A pour V1 | Maintient V1 légère | Gestion stock avancée future | Haute | À valider |
| 30 | Stock produits agricoles | Produits récoltés non vendus | Stock entrant/sortant déjà prévu | A. stock simple ; B. stock par qualité ; C. stock par lot/date | A léger ; B utile tomate/cacao ; C complexe | A en V1, champ qualité optionnel | Cohérence flux | Traçabilité future | Haute | À valider |
| 31 | Dettes fournisseurs | Dépenses restant à payer | À inclure en V1 de façon limitée | A. aucune dette ; B. suivi simple montant dû/payé ; C. échéancier détaillé | A insuffisant ; B bon compromis ; C complexe | B pour V1 | Répond besoins entrepreneur | Comptabilité future | Haute | À valider |
| 32 | Créances clients | Ventes non encaissées | À inclure en V1 de façon limitée | A. aucune créance ; B. suivi simple montant à encaisser ; C. relances/échéancier | A masque trésorerie ; B équilibré ; C hors V1 | B | Améliore résultat financier | Gestion commerciale future | Haute | À valider |
| 33 | Achats à crédit | Charge économique avec paiement différé | Distinction charge/paiement | A. ignorer ; B. dette simple ; C. échéancier complet | A faux en trésorerie ; B suffisant ; C lourd | B | Cohérence économique/financière | Finance avancée future | Critique | Recommandé |
| 34 | Paiements différés | Sorties de trésorerie à date ultérieure | Date ou période de paiement | A. pas de date ; B. mois prévu ; C. date exacte | A faible trésorerie ; B simple ; C précis | B pour V1 | Calendrier léger | Précision future | Haute | À valider |
| 35 | Avances reçues | Entrée de trésorerie non produit | Flux séparé | A. ignorer ; B. enregistrer sans lien ; C. lier à vente future | A incomplet ; B simple ; C robuste | B en V1, préparer C | Évite double comptage | Contrats futurs | Moyenne | À valider |
| 36 | Subventions | Appuis séparés du produit brut | Présentation séparée | A. exclure totalement ; B. afficher après résultat ; C. intégrer selon type | A masque appui ; B clair ; C précis mais plus lourd | B pour V1 | Performance intrinsèque visible | Analyse projets future | Moyenne | Recommandé |
| 37 | Niveau trésorerie | Calendrier de trésorerie | Flux par période | A. cycle global ; B. mois ; C. date exacte | A simple mais peu utile ; B équilibré ; C détaillé | B pour V1 | Besoin maximal crédible | Cash-flow avancé futur | Critique | À valider |
| 38 | Solde initial trésorerie | Calcul du besoin maximal | À prévoir | A. zéro par défaut ; B. saisie facultative ; C. obligatoire | A simple ; B réaliste ; C plus contraignant | B avec zéro par défaut | Améliore estimation | Finance future | Moyenne | À valider |
| 39 | Capacité remboursement | Analyse crédit | Formule simple proposée | A. non V1 ; B. ratio simple ; C. échéancier complet | A limite finance ; B utile ; C lourd | B | Utile institutions | Crédit avancé futur | Moyenne | À valider |
| 40 | Périmètre simulations | Variations prix, rendement, superficie, charges | Simulation économique demandée | A. scénario unique ; B. scénarios ±10/20 % ; C. simulateur multi-paramètres | A limité ; B simple ; C puissant mais complexe | B pour V1 | Aide décision sans surcharge | Simulation avancée future | Haute | À valider |
| 41 | Simulation charges | Variation des charges | À définir | A. variation globale ; B. par catégorie ; C. par ligne | A rapide ; B utile ; C trop détaillé | B | Analyse sensibilité | Optimisation future | Moyenne | Recommandé |
| 42 | Simulation objectifs | Objectifs production/financiers | À cadrer | A. suivi indicatif ; B. comparaison objectif/réel ; C. plan d'action détaillé | A léger ; B utile ; C lourd | B limité | Tableau de bord personnel | Pilotage avancé | Moyenne | À préciser |
| 43 | Spéculations pilotes cultures | Maïs, riz, manioc, tomate, cacao | Retenues | A. garder 5 ; B. réduire à 3 ; C. ajouter cultures | A complet ; B plus léger ; C scope creep | A, mais fiches métier par spéculation | Périmètre V1 clair | Extension paramétrable | Critique | Validé |
| 44 | Spéculations pilotes élevages | Poulet de chair, porc, pisciculture | Retenues | A. garder 3 ; B. réduire ; C. ajouter filières | A couvre diversité ; B réduit charge ; C scope creep | A avec niveau de détail limité | V1 équilibrée | Extension élevage | Critique | Validé |
| 45 | Cacao et pérennes | Analyse pluriannuelle | Prévue par phases | A. inclure cacao simple ; B. analyse phases V1 ; C. reporter pérennes | A insuffisant ; B cohérent ; C réduit périmètre | B avec modèle simplifié | Complexifie mais stratégique | VAN/TRI futurs | Haute | À préciser |
| 46 | Comptes prévisionnels/réels | Comparaison prévision/réalisation | Obligatoire V1 | A. prévisionnel seul ; B. réel seul ; C. les deux + écarts | A incomplet ; B peu planification ; C valeur forte | C | Central pour décision | Historique futur | Critique | Validé |
| 47 | Charges premier cycle | Charges propres au premier cycle | Prévu | A. ignorer ; B. champ premier cycle ; C. règle détaillée par spéculation | A faux pour cycles courts ; B simple ; C précis | B | Annualisation fiable | Paramétrage futur | Haute | À valider |
| 48 | Charges récurrentes | Charges répétées par cycle | Prévu | A. multiplier toutes charges ; B. marquer récurrente/non ; C. calendrier détaillé | A faux ; B équilibré ; C lourd | B | Calcul annuel correct | Gestion campagnes future | Haute | À valider |
| 49 | Pertes | Pertes production/post-récolte | Identifiables séparément | A. pertes globales ; B. distinguer avant/après récolte ; C. détail par cause | A simple ; B utile ; C analytique | B pour V1 | Améliore diagnostic | Analyse risque future | Haute | À valider |
| 50 | Récoltes multiples | Plusieurs récoltes par cycle | Prévu | A. total récolte ; B. lignes multiples ; C. récolte liée à vente/stock | A simple ; B nécessaire ; C robuste | B pour V1 | Tomate/pisciculture utiles | Traçabilité future | Haute | Recommandé |
| 51 | Ventes multiples | Plusieurs prix/dates | Prévu | A. prix moyen saisi ; B. lignes de vente ; C. ventes + clients + échéances | A rapide ; B fiable ; C commercial complet | B avec client facultatif | CA précis | Créances avancées | Critique | Recommandé |
| 52 | Données manquantes | Zéro vs inconnu vs non applicable | Distinction prévue | A. bloquer ; B. autoriser brouillon ; C. calcul provisoire avec IQD | A rigide ; B souple ; C utile mais à signaler | C pour brouillon, blocage validation critique | UX terrain améliorée | Moteur robuste | Critique | À valider |
| 53 | Statuts comptes | Brouillon, validé, corrigé, archivé | Prévu | A. brouillon/validé ; B. 4 statuts ; C. workflow complet | A simple ; B bon compromis ; C lourd | B | Traçabilité suffisante | Audit futur | Haute | Recommandé |
| 54 | Historique modifications | Audit des changements | Prévu | A. aucun ; B. actions critiques ; C. historique champ par champ | A risqué ; B léger ; C lourd | B en V1 | Confiance rapports | Audit complet futur | Haute | À valider |
| 55 | Capacité V1 | Volume dans Excel | À définir | A. pas de limite ; B. limite recommandée ; C. blocage technique | A risqué ; B transparent ; C contraignant | B | Performance maîtrisée | Migration SQL préparée | Haute | À préciser |
| 56 | Rapports | Présentation des résultats | Rapports imprimables | A. synthèse seule ; B. synthèse + détail ; C. rapports personnalisables | A léger ; B utile ; C lourd | B | Revue métier claire | Personnalisation future | Moyenne | À valider |
| 57 | Mode individuel | Rubriques personnelles | Liste fonctionnelle validée | A. toutes rubriques ; B. rubriques prioritaires ; C. paramétrage complet | A risque surcharge ; B V1 légère ; C lourd | B | Limite scope creep | Extension progressive | Haute | À préciser |
| 58 | Mode organisation | Consolidation | Prévue | A. simple liste résultats ; B. statistiques de base ; C. BI avancée | A faible ; B suffisant ; C hors V1 | B | Rapports institutionnels simples | BI future | Haute | À valider |
| 59 | Objectifs | Suivi objectifs production/finance | Demandé mode individuel | A. objectifs indicatifs ; B. objectifs avec écarts ; C. planification détaillée | A simple ; B utile ; C lourd | B limité | Motive entrepreneur | Pilotage avancé | Moyenne | À préciser |
| 60 | Scope creep | Stabilisation V1 | Toute extension arbitrée | A. ajouter au fil de l'eau ; B. matrice arbitrage ; C. comité formel | A dangereux ; B adapté ; C lourd | B | Protège V1 | Gouvernance durable | Critique | Recommandé |

---

# 5. Arbitrages critiques à valider en premier

Les arbitrages suivants doivent être validés avant toute conception détaillée du moteur de calcul :

| Priorité | Arbitrage | Pourquoi c'est bloquant |
|---|---|---|
| 1 | Pondérations IPE | Le classement des spéculations dépend directement de ces pondérations. |
| 2 | Pondérations IQD et seuil minimal | Les recommandations doivent tenir compte de la fiabilité des données. |
| 3 | Valorisation de la main-d'œuvre familiale | Le résultat économique dépend fortement de cette règle. |
| 4 | Valorisation de l'autoconsommation | Le produit brut économique dépend de la méthode retenue. |
| 5 | Durées d'amortissement | Les charges économiques et ROI dépendent de ces paramètres. |
| 6 | Unités locales et coefficients | Les rendements, coûts unitaires et stocks dépendent de conversions fiables. |
| 7 | Niveau de détail de la trésorerie | Le besoin maximal de trésorerie et la capacité de remboursement en dépendent. |
| 8 | Gestion des dettes et créances | Impact direct sur mode individuel, résultat financier et trésorerie. |
| 9 | Seuils d'alerte | Les diagnostics automatiques doivent être crédibles. |
| 10 | Périmètre simulations V1 | Important pour éviter le scope creep. |

---

# 6. Recommandations de gouvernance pour la revue

## 6.1 Méthode de validation proposée

Pour chaque ligne de la matrice :

1. confirmer si l'arbitrage est nécessaire en V1 ;
2. choisir une option ;
3. valider ou modifier la recommandation ;
4. attribuer un responsable métier ;
5. fixer le statut : validé, reporté ou à préciser ;
6. reporter la décision validée dans les règles de calcul ou les fiches spéculation.

## 6.2 Participants recommandés

- Propriétaire du produit.
- Expert en économie agricole.
- Conseiller agricole terrain.
- Entrepreneur agricole ou producteur structuré.
- Représentant d'organisation ou coopérative.
- Référent financier agricole, si le module trésorerie/crédit est prioritaire.

## 6.3 Règle anti-scope creep

Tout arbitrage marqué **Reporté** ou **versions futures** ne doit pas bloquer la V1. Il doit seulement être pris en compte dans l'architecture pour éviter une refonte ultérieure.

---

# 7. Synthèse des décisions à produire après revue

À la fin de la revue fonctionnelle, les décisions suivantes devront être disponibles :

- tableau final des pondérations IPE ;
- tableau final des pondérations IQD ;
- seuils d'alerte V1 ;
- barème de main-d'œuvre familiale ;
- méthode de valorisation de l'autoconsommation ;
- prix et rendements de référence par spéculation pilote ;
- durées d'amortissement par type d'actif ;
- unités locales et coefficients de conversion validés ;
- périmètre V1 de gestion des dettes, créances et trésorerie ;
- périmètre V1 des simulations ;
- fiches métier des spéculations pilotes.
