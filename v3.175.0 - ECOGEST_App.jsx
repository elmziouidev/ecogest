// Note : useState/useEffect/useRef/useCallback/useMemo sont fournis comme
// variables globales par le HTML englobant (var useState=React.useState; ...)
// — pas d'import ES module ici, car l'app tourne comme script classique.

// ════════════════════════════════════════════════════════════════════
//  SHAREPOINT CONFIG — remplacez par vos vraies valeurs Azure AD
// ════════════════════════════════════════════════════════════════════
const SP = {
  siteUrl: "https://ecolesuperieurederabat.sharepoint.com/sites/ECOGEST",
  tenantId: "72182bf9-410a-4e42-b9f0-4f740b3959e9",
  clientId: "0d52fee7-c6cf-4374-8063-395b2add8c91", // En production : authentification via MSAL (Microsoft Auth Library)
  // Les listes à créer dans SharePoint :
  lists: {
    etudiants: "ECOGEST_Etudiants",
    professeurs: "ECOGEST_Professeurs",
    notes: "ECOGEST_Notes",
    presences: "ECOGEST_Presences",
    cours: "ECOGEST_Cours",
    inscriptions: "ECOGEST_Inscriptions",
    annonces: "ECOGEST_Annonces",
    visiteurs: "ECOGEST_Visiteurs",
    filieres: "ECOGEST_Filieres",
    paiements: "ECOGEST_Paiements",
    emploiDuTemps: "ECOGEST_EmploiDuTemps",
    demandesDocuments: "ECOGEST_DemandesDocuments",
    pointageFiches: "ECOGEST_PointageFiches",
    pointageSeances: "ECOGEST_PointageSeances",
  },
};
const DEMO = {
  etudiants: [
    {
      id: 1,
      nom: "Alami",
      prenom: "Youssef",
      cne: "P123456",
      filiere: "MGE",
      annee: "2ème",
      email: "y.alami@esr.ma",
      statut: "Actif",
      genre: "M",
    },
    {
      id: 2,
      nom: "Benali",
      prenom: "Sara",
      cne: "P234567",
      filiere: "MGE",
      annee: "1ère",
      email: "s.benali@esr.ma",
      statut: "Actif",
      genre: "F",
    },
    {
      id: 3,
      nom: "Chraibi",
      prenom: "Mehdi",
      cne: "P345678",
      filiere: "MGE",
      annee: "3ème",
      email: "m.chraibi@esr.ma",
      statut: "Actif",
      genre: "M",
    },
    {
      id: 4,
      nom: "Daoudi",
      prenom: "Fatima",
      cne: "P456789",
      filiere: "MGE",
      annee: "2ème",
      email: "f.daoudi@esr.ma",
      statut: "Inactif",
      genre: "F",
    },
    {
      id: 5,
      nom: "El Idrissi",
      prenom: "Amine",
      cne: "P567890",
      filiere: "MGE",
      annee: "1ère",
      email: "a.elidrissi@esr.ma",
      statut: "Actif",
      genre: "M",
    },
  ],
  professeurs: [
    {
      id: 1,
      nom: "El Fassi",
      prenom: "Hassan",
      cin: "BE123456",
      matiere: "Finance & Comptabilité",
      email: "h.elfassi@esr.ma",
      statut: "Actif",
      type: "Permanent",
    },
    {
      id: 2,
      nom: "Tahiri",
      prenom: "Leila",
      cin: "BK234567",
      matiere: "Marketing",
      email: "l.tahiri@esr.ma",
      statut: "Actif",
      type: "Vacataire",
    },
    {
      id: 3,
      nom: "Mourad",
      prenom: "Karim",
      cin: "BB345678",
      matiere: "Droit des affaires",
      email: "k.mourad@esr.ma",
      statut: "Actif",
      type: "Permanent",
    },
    {
      id: 4,
      nom: "Soussi",
      prenom: "Nadia",
      cin: "BA456789",
      matiere: "Management stratégique",
      email: "n.soussi@esr.ma",
      statut: "Actif",
      type: "Permanent",
    },
  ],
  notes: [
    {
      id: 1,
      etudiantId: 1,
      etudiant: "Youssef Alami",
      module: "Finance",
      note: 16.5,
      semestre: "S3",
      coef: 3,
      date: "2026-06-15",
    },
    {
      id: 2,
      etudiantId: 2,
      etudiant: "Sara Benali",
      module: "Marketing",
      note: 14,
      semestre: "S1",
      coef: 2,
      date: "2026-06-15",
    },
    {
      id: 3,
      etudiantId: 3,
      etudiant: "Mehdi Chraibi",
      module: "Droit",
      note: 17,
      semestre: "S5",
      coef: 2,
      date: "2026-06-16",
    },
    {
      id: 4,
      etudiantId: 4,
      etudiant: "Fatima Daoudi",
      module: "Finance",
      note: 12,
      semestre: "S3",
      coef: 3,
      date: "2026-06-16",
    },
    {
      id: 5,
      etudiantId: 5,
      etudiant: "Amine El Idrissi",
      module: "Marketing",
      note: 15,
      semestre: "S1",
      coef: 2,
      date: "2026-06-17",
    },
  ],
  presences: [
    { id: 1, etudiant: "Youssef Alami", module: "Finance", date: "2026-07-20", heure: "08h00", statut: "Présent" },
    { id: 2, etudiant: "Sara Benali", module: "Marketing", date: "2026-07-20", heure: "10h00", statut: "Absent" },
    { id: 3, etudiant: "Mehdi Chraibi", module: "Droit", date: "2026-07-21", heure: "08h00", statut: "Présent" },
    { id: 4, etudiant: "Fatima Daoudi", module: "Finance", date: "2026-07-21", heure: "08h00", statut: "Justifié" },
    { id: 5, etudiant: "Amine El Idrissi", module: "Marketing", date: "2026-07-21", heure: "10h00", statut: "Présent" },
  ],
  cours: [
    {
      id: 1,
      intitule: "Finance d'entreprise",
      professeur: "Hassan El Fassi",
      filiere: "MGE",
      semestre: "S3",
      heures: 40,
      salle: "A101",
    },
    {
      id: 2,
      intitule: "Marketing opérationnel",
      professeur: "Leila Tahiri",
      filiere: "MGE",
      semestre: "S3",
      heures: 30,
      salle: "B202",
    },
    {
      id: 3,
      intitule: "Droit des sociétés",
      professeur: "Karim Mourad",
      filiere: "MGE",
      semestre: "S3",
      heures: 30,
      salle: "A102",
    },
    {
      id: 4,
      intitule: "Management stratégique",
      professeur: "Nadia Soussi",
      filiere: "MGE",
      semestre: "S5",
      heures: 40,
      salle: "C301",
    },
  ],
  annonces: [
    {
      id: 1,
      titre: "Rentrée 2026-2027",
      contenu: "La rentrée est fixée au 15 septembre 2026.",
      auteur: "Direction",
      date: "2026-07-01",
      cible: "Tous",
    },
    {
      id: 2,
      titre: "Dépôt des PFA",
      contenu: "Les projets de fin d'année doivent être déposés avant le 30 juin.",
      auteur: "Scolarité",
      date: "2026-06-01",
      cible: "Étudiants",
    },
  ],
  visiteurs: [
    {
      id: 1,
      nom: "Rachidi",
      prenom: "Imane",
      cin: "BK123456",
      dateNaissance: "2005-03-14",
      lieuNaissance: "Rabat",
      telephone: "0612345678",
      email: "i.rachidi@gmail.com",
      adresse: "12 Rue des Fleurs",
      ville: "Rabat",
      quartier: "Agdal",
      bacAnnee: "2024",
      bacMention: "Bien",
      bacSerie: "Sciences Économiques",
      etablissement: "Lycée Descartes",
      filiereS: "BAC+3 en Management et Gestion des entreprises",
      dateVisite: "2026-07-20",
      statut: "Visiteur",
      tarif: [
        { tranche: "1ère tranche", montant: 6e3, echeance: "2026-09-15", paye: false },
        { tranche: "2ème tranche", montant: 6e3, echeance: "2027-01-15", paye: false },
        { tranche: "3ème tranche", montant: 6e3, echeance: "2027-04-15", paye: false },
      ],
      documents: [],
    },
    {
      id: 2,
      nom: "Benkirane",
      prenom: "Karim",
      cin: "BE234567",
      dateNaissance: "2004-07-22",
      lieuNaissance: "Salé",
      telephone: "0623456789",
      email: "k.benkirane@gmail.com",
      adresse: "5 Avenue Mohammed V",
      ville: "Salé",
      quartier: "Centre",
      bacAnnee: "2023",
      bacMention: "Passable",
      bacSerie: "Sciences de Gestion",
      etablissement: "Lycée Ibn Khaldoun",
      filiereS: "BAC+3 en Systèmes d'Information – Option Génie Logiciel",
      dateVisite: "2026-07-21",
      statut: "Inscrit",
      tarif: [
        { tranche: "1ère tranche", montant: 6500, echeance: "2026-09-15", paye: true },
        { tranche: "2ème tranche", montant: 6500, echeance: "2027-01-15", paye: false },
        { tranche: "3ème tranche", montant: 6500, echeance: "2027-04-15", paye: false },
      ],
      documents: [],
    },
  ],
  paiements: null,
  filieres: [
    {
      id: 3,
      code: "MGE",
      intitule: "BAC+3 en Management et Gestion des entreprises",
      niveau: "BAC+3",
      duree: 3,
      capacite: 35,
      frais: 18e3,
      statut: "Active",
      description: "Formation en management, gestion et finance d'entreprise.",
      responsable: "",
      couleur: "#e56b2d",
    },
    {
      id: 4,
      code: "ASI",
      intitule: "BAC+5 en Audit des Systèmes d'Information",
      niveau: "BAC+5",
      duree: 2,
      capacite: 25,
      frais: 24e3,
      statut: "Active",
      description: "Formation en audit et contrôle des systèmes d'information.",
      responsable: "",
      couleur: "#8b5cf6",
    },
    {
      id: 5,
      code: "GLM",
      intitule: "BAC+5 en Génie Logiciel et Multimédia",
      niveau: "BAC+5",
      duree: 2,
      capacite: 25,
      frais: 24e3,
      statut: "Active",
      description: "Formation avancée en développement logiciel et multimédia.",
      responsable: "",
      couleur: "#10b981",
    },
    {
      id: 6,
      code: "CRSI",
      intitule: "BAC+5 en Cybersécurité des réseaux et des systèmes d'information",
      niveau: "BAC+5",
      duree: 2,
      capacite: 25,
      frais: 25e3,
      statut: "Active",
      description: "Formation en sécurité informatique et cyberdéfense.",
      responsable: "",
      couleur: "#ef4444",
    },
    {
      id: 7,
      code: "MSMD",
      intitule: "BAC+5 en Marketing Stratégique et Marketing Digital",
      niveau: "BAC+5",
      duree: 2,
      capacite: 25,
      frais: 23e3,
      statut: "Active",
      description: "Formation en marketing stratégique et digital.",
      responsable: "",
      couleur: "#f59e0b",
    },
    {
      id: 8,
      code: "MS",
      intitule: "BAC+5 en Management Stratégique",
      niveau: "BAC+5",
      duree: 2,
      capacite: 25,
      frais: 23e3,
      statut: "Active",
      description: "Formation en management et stratégie d'entreprise.",
      responsable: "",
      couleur: "#1a2f5e",
    },
  ],
};
const C = {
  navy: "#1a2f5e",
  orange: "#e56b2d",
  green: "#2e7d52",
  red: "#c0392b",
  purple: "#7c3aed",
  teal: "#0d9488",
  light: "#f6f7f9",
  border: "#e8eaef",
  muted: "#888f9c",
  white: "#ffffff",
  navyLight: "#eef1f8",
  greenLight: "#e9f6ee",
  orangeLight: "#fdf1e9", // Nouveaux jetons de design (ombres, rayons) — utilisés pour donner plus de
  // profondeur/air aux cartes, boutons et modales sans changer les couleurs
  // de marque (logo, PDF) qui restent identiques.
  shadowSm: "0 1px 2px rgba(16,24,50,0.05), 0 1px 1px rgba(16,24,50,0.04)",
  shadow: "0 2px 6px rgba(16,24,50,0.06), 0 8px 24px rgba(16,24,50,0.06)",
  shadowLg: "0 24px 60px rgba(10,16,35,0.22)",
  radius: 14,
  radiusSm: 10,
};
const ROLES = {
  directrice: { label: "Directrice", color: C.purple, icon: "👩‍💼" },
  administrateur: { label: "Administrateur", color: C.navy, icon: "⚙️" },
  professeur: { label: "Professeur", color: C.green, icon: "👨‍🏫" },
  etudiant: { label: "Étudiant", color: C.orange, icon: "🎓" },
};
const mention = (n) =>
  n >= 16
    ? ["Très bien", C.green]
    : n >= 14
      ? ["Bien", C.navy]
      : n >= 10
        ? ["Passable", C.orange]
        : ["Insuffisant", C.red];
const filiereCode = (nomFiliere, listeFilieres) => {
  if (!nomFiliere) return "";
  const match = (listeFilieres || []).find((f) => f.intitule === nomFiliere || f.code === nomFiliere);
  return match?.code || nomFiliere;
};
const ETUDIANTS_TEST = ["ACHRAF HAKIMI", "HANI CHAKER"];
const estEtudiantTest = (prenom, nom) => ETUDIANTS_TEST.includes(`${prenom || ""} ${nom || ""}`.trim().toUpperCase());
const PALETTE_FILIERES = [
  "#e56b2d",
  "#0e9187",
  "#7c3aed",
  "#2563eb",
  "#c2185b",
  "#059669",
  "#d97706",
  "#0891b2",
  "#be123c",
  "#4f46e5",
  "#65a30d",
  "#9333ea",
];
const colorForFiliereCode = (code) => {
  const s = String(code || "");
  let hash = 0;
  for (let i = 0; i < s.length; i++) hash = (hash * 31 + s.charCodeAt(i)) >>> 0;
  return PALETTE_FILIERES[hash % PALETTE_FILIERES.length];
};
const anneeDeFiliere = (intitule) => (intitule || "").split(",")[0].trim();
const masquerCinProf = (nomComplet) => (nomComplet || "").replace(/\s*\(CIN\s*:\s*[^)]*\)/i, "").trim();
const colorForAnnee = (intitule) => colorForFiliereCode(anneeDeFiliere(intitule));
const FiliereBadge = ({ code }) => (
  /* @__PURE__ */ <span
    style={{
      background: colorForFiliereCode(code),
      color: "#ffffff",
      fontWeight: 800,
      fontSize: 11.5,
      padding: "4px 11px",
      borderRadius: 20,
      boxShadow: `0 2px 6px ${colorForFiliereCode(code)}66`,
      display: "inline-block",
      letterSpacing: "0.02em",
      whiteSpace: "nowrap",
    }}
  >
    {code || "—"}
  </span>
);
const Badge = ({ children, color = C.orange, size = 12 }) => (
  /* @__PURE__ */ <span
    style={{
      background: color + "16",
      color,
      border: `1px solid ${color}35`,
      borderRadius: 20,
      padding: "3px 10px",
      fontSize: size,
      fontWeight: 600,
      whiteSpace: "nowrap",
      letterSpacing: "0.01em",
    }}
  >
    {children}
  </span>
);
const IconBtn = ({ icon, title, color = C.navy, onClick, disabled }) => (
  /* @__PURE__ */ <button
    title={title}
    onClick={onClick}
    disabled={disabled}
    style={{
      border: "none",
      borderRadius: 8,
      cursor: disabled ? "not-allowed" : "pointer",
      width: 32,
      height: 32,
      minWidth: 32,
      fontSize: 14,
      background: disabled ? C.light : color,
      color: disabled ? "#999" : "white",
      opacity: disabled ? 0.6 : 1,
      display: "flex",
      alignItems: "center",
      justifyContent: "center",
      flexShrink: 0,
      transition: "filter 0.12s",
    }}
    onMouseEnter={(e) => {
      if (!disabled) e.currentTarget.style.filter = "brightness(1.1)";
    }}
    onMouseLeave={(e) => (e.currentTarget.style.filter = "brightness(1)")}
  >
    {icon}
  </button>
);
const BTN_GRADIENTS = {
  primary: `linear-gradient(135deg, ${C.navy} 0%, ${C.teal} 100%)`,
  navy: `linear-gradient(135deg, ${C.navy} 0%, ${C.teal} 100%)`,
  danger: `linear-gradient(135deg, ${C.red} 0%, #8e2418 100%)`,
};
const BTN_GLOW = { primary: "rgba(26,47,94,0.4)", navy: "rgba(26,47,94,0.4)", danger: "rgba(192,57,43,0.4)" };
const Btn = ({ children, onClick, variant = "primary", small, disabled, style = {} }) => (
  /* @__PURE__ */ <button
    onClick={onClick}
    disabled={disabled}
    onMouseEnter={(e) => {
      if (!disabled) {
        e.currentTarget.style.transform = "translateY(-2px)";
        e.currentTarget.style.filter = "brightness(1.06)";
      }
    }}
    onMouseLeave={(e) => {
      e.currentTarget.style.transform = "translateY(0)";
      e.currentTarget.style.filter = "brightness(1)";
    }}
    style={{
      padding: small ? "7px 16px" : "10px 22px",
      borderRadius: 12,
      border: variant === "light" ? `1px solid ${C.border}` : "none",
      cursor: disabled ? "not-allowed" : "pointer",
      fontSize: small ? 12 : 13.5,
      fontWeight: 700,
      fontFamily: "inherit",
      opacity: disabled ? 0.5 : 1,
      background: variant === "light" ? C.white : BTN_GRADIENTS[variant] || C.light,
      color: variant === "light" ? "#3d4356" : C.white,
      boxShadow: variant === "light" ? "none" : `0 4px 12px ${BTN_GLOW[variant] || "rgba(0,0,0,0.14)"}`,
      transition: "transform 0.15s ease, box-shadow 0.15s ease, filter 0.15s ease, opacity 0.15s",
      ...style,
    }}
  >
    {children}
  </button>
);
const Input = ({ label, ...p }) => (
  /* @__PURE__ */ <div style={{ marginBottom: 13 }}>
    {label && (
      /* @__PURE__ */ <label
        style={{ fontSize: 12, fontWeight: 600, color: "#555", display: "block", marginBottom: 4 }}
      >
        {label}
      </label>
    )}
    <input
      {...p}
      style={{
        width: "100%",
        padding: "9px 12px",
        border: `1px solid ${C.border}`,
        borderRadius: 9,
        fontSize: 13,
        fontFamily: "inherit",
        outline: "none",
        boxSizing: "border-box",
        transition: "border-color 0.12s",
        ...p.style,
      }}
      onFocus={(e) => {
        e.currentTarget.style.borderColor = C.navy;
        p.onFocus && p.onFocus(e);
      }}
      onBlur={(e) => {
        e.currentTarget.style.borderColor = C.border;
        p.onBlur && p.onBlur(e);
      }}
    />
  </div>
);
const Sel = ({ label, children, ...p }) => (
  /* @__PURE__ */ <div style={{ marginBottom: 13 }}>
    {label && (
      /* @__PURE__ */ <label
        style={{ fontSize: 12, fontWeight: 600, color: "#555", display: "block", marginBottom: 4 }}
      >
        {label}
      </label>
    )}
    <select
      {...p}
      style={{
        width: "100%",
        padding: "9px 12px",
        border: `1px solid ${C.border}`,
        borderRadius: 9,
        fontSize: 13,
        fontFamily: "inherit",
        outline: "none",
        boxSizing: "border-box",
        background: C.white,
      }}
    >
      {children}
    </select>
  </div>
);
const Modal = ({ title, onClose, children, width = 480, hideClose, hideFooterClose }) => (
  /* @__PURE__ */ <div
    style={{
      position: "fixed",
      inset: 0,
      background: "rgba(10,16,35,0.5)",
      backdropFilter: "blur(2px)",
      zIndex: 1e3,
      display: "flex",
      alignItems: "center",
      justifyContent: "center",
    }}
  >
    <div
      className="eco-modal"
      style={{
        background: C.white,
        borderRadius: 16,
        padding: 28,
        width,
        maxWidth: "92vw",
        boxShadow: C.shadowLg,
        maxHeight: "90vh",
        overflowY: "auto",
      }}
    >
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 20 }}>
        <h3 style={{ margin: 0, color: C.navy, fontSize: 17, fontWeight: 700, letterSpacing: "-0.01em" }}>{title}</h3>
        {!hideClose && (
          /* @__PURE__ */ <button
            onClick={onClose}
            style={{
              background: C.light,
              border: "none",
              width: 28,
              height: 28,
              borderRadius: 8,
              fontSize: 15,
              cursor: "pointer",
              color: "#888",
              lineHeight: 1,
            }}
          >
            ✕
          </button>
        )}
      </div>
      {children}
      {!hideClose && !hideFooterClose && (
        /* @__PURE__ */ <div
          style={{
            display: "flex",
            justifyContent: "flex-end",
            marginTop: 18,
            paddingTop: 14,
            borderTop: `1px solid ${C.light}`,
          }}
        >
          <Btn variant="light" onClick={onClose}>
            Fermer
          </Btn>
        </div>
      )}
    </div>
  </div>
);
const DataTable = ({ cols, rows, emptyMsg = "Aucune donnée", onRowClick, compact, rowStyle }) => (
  /* @__PURE__ */ <div style={{ overflowX: "auto", border: `1px solid ${C.border}`, borderRadius: 12 }}>
    <table style={{ width: "100%", borderCollapse: "collapse", fontSize: compact ? 12 : 13 }}>
      <thead>
        <tr style={{ background: `linear-gradient(135deg, ${C.navy} 0%, ${C.teal} 100%)` }}>
          {cols.map((c, i) => (
            /* @__PURE__ */ <th
              key={i}
              style={{
                padding: compact ? "7px 12px" : "11px 14px",
                textAlign: c.headerAlign || "left",
                fontWeight: 700,
                color: "rgba(255,255,255,0.92)",
                fontSize: compact ? 9.5 : 10.5,
                textTransform: "uppercase",
                letterSpacing: "0.06em",
                whiteSpace: "nowrap",
                background: c.headerBg || "transparent",
              }}
            >
              {c.label}
            </th>
          ))}
        </tr>
      </thead>
      <tbody>
        {rows.length === 0 ? (
          /* @__PURE__ */ <tr>
            <td
              colSpan={cols.length}
              style={{ padding: compact ? 20 : 32, textAlign: "center", color: "#bbb", fontSize: 13 }}
            >
              {emptyMsg}
            </td>
          </tr>
        ) : (
          rows.map((row, i) => {
            const rs = rowStyle && rowStyle(row, i);
            const isObj = rs && typeof rs === "object";
            const baseBg = (isObj ? rs.background : rs) || (i % 2 === 0 ? C.white : C.greenLight);
            return (
              /* @__PURE__ */ <tr
                key={i}
                onClick={onRowClick ? () => onRowClick(row) : void 0}
                style={{
                  background: baseBg,
                  borderBottom: (isObj && rs.borderBottom) || `1px solid ${C.border}`,
                  borderTop: isObj ? rs.borderTop : void 0,
                  cursor: onRowClick ? "pointer" : "default",
                  transition: "background 0.12s",
                }}
                onMouseEnter={(e) => (e.currentTarget.style.background = C.orangeLight)}
                onMouseLeave={(e) => (e.currentTarget.style.background = baseBg)}
              >
                {cols.map((c, j) => (
                  /* @__PURE__ */ <td key={j} style={{ padding: compact ? "5px 12px" : "11px 14px", color: "#333" }}>
                    {c.render ? c.render(row[c.key], row) : row[c.key]}
                  </td>
                ))}
              </tr>
            );
          })
        )}
      </tbody>
    </table>
  </div>
);
const Card = ({ children, style = {} }) => (
  /* @__PURE__ */ <div
    style={{
      background: C.white,
      border: `1px solid ${C.border}`,
      borderRadius: C.radius,
      boxShadow: C.shadowSm,
      ...style,
    }}
  >
    {children}
  </div>
);
const StatCard = ({ label, value, color, icon }) => (
  /* @__PURE__ */ <div
    style={{
      background: C.white,
      border: `1px solid ${C.border}`,
      borderRadius: C.radius,
      padding: "16px 20px",
      flex: 1,
      minWidth: 116,
      boxShadow: C.shadowSm,
    }}
  >
    <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 10 }}>
      <div
        style={{
          width: 26,
          height: 26,
          borderRadius: 8,
          background: color + "16",
          display: "flex",
          alignItems: "center",
          justifyContent: "center",
          fontSize: 13,
          flexShrink: 0,
        }}
      >
        {icon}
      </div>
      <div style={{ fontSize: 11, color: C.muted, fontWeight: 600, letterSpacing: "0.01em" }}>{label}</div>
    </div>
    <div style={{ fontSize: 26, fontWeight: 800, color, letterSpacing: "-0.02em" }}>{value}</div>
  </div>
);
const PageHeader = ({ title, sub, action }) => (
  /* @__PURE__ */ <div
    className="eco-page-header"
    style={{
      display: "flex",
      justifyContent: "space-between",
      alignItems: "flex-start",
      flexWrap: "wrap",
      gap: 14,
      marginBottom: 22,
    }}
  >
    <div style={{ minWidth: 0 }}>
      <h2 style={{ color: C.navy, margin: 0, fontSize: 21, fontWeight: 800, letterSpacing: "-0.015em" }}>{title}</h2>
      {sub && /* @__PURE__ */ <p style={{ color: C.muted, fontSize: 13, margin: "5px 0 0" }}>{sub}</p>}
    </div>
    {action}
  </div>
);
const Toast = ({ msg }) =>
  msg ? (
    /* @__PURE__ */ <div
      style={{
        position: "fixed",
        bottom: 24,
        right: 24,
        background: C.navy,
        color: C.white,
        padding: "12px 20px",
        borderRadius: 11,
        fontSize: 13,
        fontWeight: 500,
        zIndex: 2e3,
        boxShadow: C.shadowLg,
      }}
    >
      ✓ {msg}
    </div>
  ) : null;
function Dashboard({ data, setData, role, user, toast }) {
  const paiementsExistants = data.paiements?.etudiants_paiement || [];
  let totalAPayerGlobal = 0,
    totalPayeGlobal = 0;
  (data.etudiants || []).forEach((et) => {
    const existing = paiementsExistants.find(
      (e) => (et.numInscription && e.numInscription === et.numInscription) || e.etudiantId === et.id
    );
    if (existing) {
      totalAPayerGlobal += existing.totalAPayer || 0;
      totalPayeGlobal += existing.totalPaye || 0;
    } else {
      const totalConvention = et.tarifGlobalFormation || 0;
      totalAPayerGlobal += totalConvention;
    }
  });
  const tauxRecouvrement = totalAPayerGlobal > 0 ? Math.round((totalPayeGlobal / totalAPayerGlobal) * 100) : 0;
  const derniersVisiteurs = [...(data.visiteurs || [])]
    .sort((a, b) => (b.dateVisite || "").localeCompare(a.dateVisite || "") || (b.id || 0) - (a.id || 0))
    .slice(0, 4);
  const tousVersements = (data.paiements?.etudiants_paiement || []).flatMap((ep) =>
    (ep.versements || []).map((v) => ({ ...v, etudiantNom: `${ep.prenom || ""} ${ep.nom || ""}`.trim() }))
  );
  const derniersPaiements = [...tousVersements].sort((a, b) => (b.id || 0) - (a.id || 0)).slice(0, 4);
  const dernieresAbsences = [...(data.presences || [])]
    .filter((p) => p.statut === "Absent")
    .sort((a, b) => (b.date || "").localeCompare(a.date || "") || (b.id || 0) - (a.id || 0))
    .slice(0, 4);
  const dernieresEvaluations = [...(data.notes || [])]
    .sort((a, b) => (b.date || "").localeCompare(a.date || "") || (b.id || 0) - (a.id || 0))
    .slice(0, 4);
  const derniersEtudiants = [...(data.etudiants || [])]
    .sort((a, b) => (parseInt(b.id) || 0) - (parseInt(a.id) || 0))
    .slice(0, 4);
  return (
    /* @__PURE__ */ <div>
      <PageHeader title="Tableau de bord" sub={`Bienvenue sur ECOGEST — École Supérieure de Rabat`} />
      <div style={{ display: "flex", gap: 12, flexWrap: "wrap", marginBottom: 24 }}>
        <StatCard
          label="Étudiants actifs"
          value={data.etudiants.filter((e) => e.statut === "Actif").length}
          color={C.navy}
          icon="🎓"
        />
        <StatCard label="Groupes" value={(data.filieres || []).length} color={C.teal} icon="🎯" />
        <StatCard label="Visiteurs" value={(data.visiteurs || []).length} color={C.orange} icon="👋" />
        <StatCard label="Total encaissé" value={formatMAD(totalPayeGlobal)} color={C.green} icon="💵" />
        <StatCard label="Taux de recouvrement" value={`${tauxRecouvrement}%`} color={C.purple} icon="📈" />
      </div>
      <div className="eco-grid2" style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16, marginBottom: 16 }}>
        <Card style={{ padding: 20 }}>
          <h4 style={{ margin: "0 0 14px", color: C.navy, fontSize: 14 }}>👋 Derniers visiteurs passés</h4>
          {derniersVisiteurs.length === 0 ? (
            /* @__PURE__ */ <div style={{ textAlign: "center", padding: 16, color: C.muted, fontSize: 12.5 }}>
              Aucun visiteur enregistré
            </div>
          ) : (
            derniersVisiteurs.map((v) => (
              /* @__PURE__ */ <div
                key={v.id}
                style={{
                  display: "flex",
                  justifyContent: "space-between",
                  alignItems: "center",
                  padding: "8px 0",
                  borderBottom: `1px solid ${C.light}`,
                }}
              >
                <div>
                  <div style={{ fontWeight: 600, fontSize: 13 }}>
                    {v.prenom} {v.nom}
                  </div>
                  <div style={{ fontSize: 11, color: C.muted }}>
                    {v.filiereS || "—"} · {v.dateVisite || "—"}
                  </div>
                </div>
                <Badge color={v.statut === "Inscrit" ? C.green : v.statut === "Non intéressé" ? C.red : C.orange}>
                  {v.statut}
                </Badge>
              </div>
            ))
          )}
        </Card>
        <Card style={{ padding: 20 }}>
          <h4 style={{ margin: "0 0 14px", color: C.navy, fontSize: 14 }}>🎓 Derniers étudiants inscrits</h4>
          {derniersEtudiants.map((e) => (
            /* @__PURE__ */ <div
              key={e.id}
              style={{
                display: "flex",
                justifyContent: "space-between",
                alignItems: "center",
                padding: "8px 0",
                borderBottom: `1px solid ${C.light}`,
              }}
            >
              <div>
                <div style={{ fontWeight: 600, fontSize: 13 }}>
                  {e.prenom} {e.nom}
                </div>
                <div style={{ fontSize: 11, color: C.muted }}>
                  {e.filiere} · {e.email}
                </div>
              </div>
              <Badge color={e.statut === "Actif" ? C.green : C.muted}>{e.statut}</Badge>
            </div>
          ))}
        </Card>
        <Card style={{ padding: 20 }}>
          <h4 style={{ margin: "0 0 14px", color: C.navy, fontSize: 14 }}>💰 Derniers paiements faits</h4>
          {derniersPaiements.length === 0 ? (
            /* @__PURE__ */ <div style={{ textAlign: "center", padding: 16, color: C.muted, fontSize: 12.5 }}>
              Aucun paiement enregistré
            </div>
          ) : (
            derniersPaiements.map((v, i) => (
              /* @__PURE__ */ <div
                key={v.id || i}
                style={{
                  display: "flex",
                  justifyContent: "space-between",
                  alignItems: "center",
                  padding: "8px 0",
                  borderBottom: `1px solid ${C.light}`,
                }}
              >
                <div>
                  <div style={{ fontWeight: 600, fontSize: 13 }}>{v.etudiantNom}</div>
                  <div style={{ fontSize: 11, color: C.muted }}>
                    {v.mode} · {formatDateFR(v.date)}
                  </div>
                </div>
                <div style={{ fontWeight: 700, color: C.green, fontSize: 14 }}>{formatMAD(v.montant)}</div>
              </div>
            ))
          )}
        </Card>
        <Card style={{ padding: 20 }}>
          <h4 style={{ margin: "0 0 14px", color: C.navy, fontSize: 14 }}>❌ Dernières absences enregistrées</h4>
          {dernieresAbsences.length === 0 ? (
            /* @__PURE__ */ <div style={{ textAlign: "center", padding: 16, color: C.muted, fontSize: 12.5 }}>
              Aucune absence enregistrée
            </div>
          ) : (
            dernieresAbsences.map((p, i) => (
              /* @__PURE__ */ <div
                key={p.id || i}
                style={{
                  display: "flex",
                  justifyContent: "space-between",
                  alignItems: "center",
                  padding: "8px 0",
                  borderBottom: `1px solid ${C.light}`,
                }}
              >
                <div>
                  <div style={{ fontWeight: 600, fontSize: 13 }}>{p.etudiant}</div>
                  <div style={{ fontSize: 11, color: C.muted }}>
                    {p.module} · {formatDateFR(p.date)}
                  </div>
                </div>
                <Badge color={C.red}>{p.statut}</Badge>
              </div>
            ))
          )}
        </Card>
      </div>
      <Card style={{ padding: 20, marginBottom: 16 }}>
        <h4 style={{ margin: "0 0 14px", color: C.navy, fontSize: 14 }}>📊 Dernières évaluations passées</h4>
        {dernieresEvaluations.length === 0 ? (
          /* @__PURE__ */ <div style={{ textAlign: "center", padding: 16, color: C.muted, fontSize: 12.5 }}>
            Aucune évaluation enregistrée
          </div>
        ) : (
          /* @__PURE__ */ <div
            className="eco-grid2"
            style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "0 24px" }}
          >
            {dernieresEvaluations.map((n, i) => {
              const [men, col] = mention(n.note);
              return (
                /* @__PURE__ */ <div
                  key={n.id || i}
                  style={{
                    display: "flex",
                    justifyContent: "space-between",
                    alignItems: "center",
                    padding: "8px 0",
                    borderBottom: `1px solid ${C.light}`,
                  }}
                >
                  <div>
                    <div style={{ fontWeight: 600, fontSize: 13 }}>{n.etudiant}</div>
                    <div style={{ fontSize: 11, color: C.muted }}>
                      {n.module} — {n.semestre} · {formatDateFR(n.date)}
                    </div>
                  </div>
                  <div style={{ textAlign: "right" }}>
                    <div style={{ fontWeight: 700, color: col, fontSize: 15 }}>{n.note}/20</div>
                    <Badge color={col} size={10}>
                      {men}
                    </Badge>
                  </div>
                </div>
              );
            })}
          </div>
        )}
      </Card>
      <Card style={{ padding: 20 }}>
        <h4 style={{ margin: "0 0 14px", color: C.navy, fontSize: 14 }}>🗓️ Calendrier — Prochains événements</h4>
        {(() => {
          const prochains = [...(data.calendrier || [])]
            .filter((e) => !e.dateDebut || e.dateDebut >= today())
            .sort((a, b) => (a.dateDebut || "").localeCompare(b.dateDebut || ""))
            .slice(0, 4);
          if (prochains.length === 0)
            return (
              /* @__PURE__ */ <div style={{ textAlign: "center", padding: 16, color: C.muted, fontSize: 12.5 }}>
                Aucun événement à venir
              </div>
            );
          return prochains.map((ev) => {
            const t = TYPES_EVENEMENT[ev.type] || TYPES_EVENEMENT.Autre;
            return (
              /* @__PURE__ */ <div
                key={ev.id}
                style={{ display: "flex", gap: 14, padding: "10px 0", borderBottom: `1px solid ${C.light}` }}
              >
                <div
                  style={{
                    width: 36,
                    height: 36,
                    borderRadius: 8,
                    background: t.color + "18",
                    display: "flex",
                    alignItems: "center",
                    justifyContent: "center",
                    fontSize: 18,
                    flexShrink: 0,
                  }}
                >
                  {t.icon}
                </div>
                <div>
                  <div style={{ fontWeight: 600, fontSize: 13, color: C.navy }}>{ev.titre}</div>
                  <div style={{ fontSize: 11, color: C.muted, marginTop: 2 }}>
                    {formatDateFR(ev.dateDebut)}
                    {ev.dateFin && ev.dateFin !== ev.dateDebut ? ` → ${formatDateFR(ev.dateFin)}` : ""} · <Badge size={10} color={t.color}>
                      {ev.type}
                    </Badge>
                  </div>
                </div>
              </div>
            );
          });
        })()}
      </Card>
    </div>
  );
}
function Etudiants({ data, setData, role, user, toast }) {
  const canEdit = ["administrateur", "directrice"].includes(role);
  const [search, setSearch] = useState("");
  const [filterFiliere, setFilterFiliere] = useState("");
  const [filterStatut, setFilterStatut] = useState("");
  const [showModal, setShowModal] = useState(false);
  const [editItem, setEditItem] = useState(null);
  const emptyForm = {
    nom: "",
    prenom: "",
    cin: "",
    dateNaissance: "",
    lieuNaissance: "",
    filiere: "",
    niveau: "BAC+3",
    annee: "",
    email: "",
    emailPersonnel: "",
    telephone: "",
    statut: "Actif",
    genre: "M",
    ville: "",
    quartier: "",
    adresse: "",
    numInscription: "",
    promotion: "",
    responsableFinancier: "",
    inscritPar: "",
    bacAnnee: "",
    bacMention: "Passable",
    bacSerie: "",
    etablissement: "",
    observations: "",
    tarifGlobalFormation: 0,
    modeReglement: "Trimestriel",
    notesConvention: "",
  };
  const [form, setForm] = useState(emptyForm);
  const [doc, setDoc] = useState([]);
  const [uploading, setUploading] = useState(false);
  const [loadingDocs, setLoadingDocs] = useState(false);
  const [docCounts, setDocCounts] = useState({});
  React.useEffect(() => {
    setDoc([]);
    if (!editItem || !user?.token) return;
    setLoadingDocs(true);
    loadDocs(user.token).finally(() => setLoadingDocs(false));
  }, [editItem?.id]);
  const getDriveId = async (token) => {
    if (window._spDriveId) return window._spDriveId;
    try {
      const siteId = await getSiteIdCached(token);
      if (!siteId) {
        console.error("getDriveId: pas de siteId");
        return null;
      }
      const res = await fetch(`https://graph.microsoft.com/v1.0/sites/${siteId}/drives`, {
        headers: { Authorization: `Bearer ${token}` },
      });
      const data2 = await res.json();
      if (data2.error) {
        console.error("Drive API error:", data2.error.message);
        return null;
      }
      if (data2.value && data2.value.length > 0) {
        const docDrive =
          data2.value.find(
            (d) =>
              d.name === "Documents" ||
              d.name === "Documents partages" ||
              d.name === "Documents partagés" ||
              d.driveType === "documentLibrary"
          ) || data2.value[0];
        window._spDriveId = docDrive.id;
        return docDrive.id;
      }
    } catch (e) {
      console.error("getDriveId exception:", e.message);
    }
    return null;
  };
  const etudiantFolderName = (et) =>
    [et?.cin || "", et?.nom || "", et?.prenom || ""]
      .filter(Boolean)
      .join("_")
      .replace(/[^a-zA-Z0-9_\u00C0-\u017E]/g, "_")
      .slice(0, 60) || "divers";
  const openEtudiantFolder = async (et) => {
    if (!user?.token) {
      toast("⚠️ Non connecté à SharePoint");
      return;
    }
    try {
      const driveId = await getDriveId(user.token);
      if (!driveId) {
        toast("⚠️ Impossible de se connecter au stockage SharePoint");
        return;
      }
      const folder = `ECOGEST_Etudiants/${etudiantFolderName(et)}`;
      const res = await fetch(
        `https://graph.microsoft.com/v1.0/drives/${driveId}/root:/${encodeURIComponent(folder).replace(/%2F/g, "/")}`,
        { headers: { Authorization: `Bearer ${user.token}` } }
      );
      const d = await res.json();
      if (d.webUrl) {
        window.open(d.webUrl, "_blank");
      } else {
        toast("📁 Aucun dossier trouvé — aucun document n'a encore été déposé pour cet étudiant");
      }
    } catch (e) {
      console.error("openEtudiantFolder error:", e.message);
      toast("⚠️ Erreur lors de l'ouverture du dossier");
    }
  };
  React.useEffect(() => {
    if (!user?.token || !data.etudiants?.length) return;
    let annule = false;
    (async () => {
      const driveId = await getDriveId(user.token);
      if (!driveId || annule) return;
      const resultats = await Promise.all(
        data.etudiants.map(async (et) => {
          try {
            const folder = `ECOGEST_Etudiants/${etudiantFolderName(et)}`;
            const res = await fetch(
              `https://graph.microsoft.com/v1.0/drives/${driveId}/root:/${encodeURIComponent(folder).replace(/%2F/g, "/")}`,
              { headers: { Authorization: `Bearer ${user.token}` } }
            );
            const d = await res.json();
            return [et.id, d?.folder?.childCount ?? 0];
          } catch {
            return [et.id, 0];
          }
        })
      );
      if (!annule) setDocCounts(Object.fromEntries(resultats));
    })();
    return () => {
      annule = true;
    };
  }, [data.etudiants, user?.token]);
  const handleUpload = async (e, type) => {
    const file = e.target.files[0];
    if (!file || !user?.token || !editItem) return;
    setUploading(true);
    const MAX_SIMPLE = 4 * 1024 * 1024;
    const CHUNK_SIZE = 5 * 1024 * 1024;
    try {
      const driveId = await getDriveId(user.token);
      if (!driveId) {
        alert("Impossible de se connecter au stockage SharePoint.");
        setUploading(false);
        return;
      }
      const folder = `ECOGEST_Etudiants/${etudiantFolderName(editItem)}`;
      const fileName = type.replace(/[^a-zA-Z0-9]/g, "_") + "_" + file.name;
      const filePath = `${folder}/${fileName}`;
      let resultData;
      if (file.size <= MAX_SIMPLE) {
        const ab = await file.arrayBuffer();
        const res = await fetch(`https://graph.microsoft.com/v1.0/drives/${driveId}/root:/${filePath}:/content`, {
          method: "PUT",
          headers: { Authorization: `Bearer ${user.token}`, "Content-Type": file.type || "application/octet-stream" },
          body: ab,
        });
        resultData = await res.json();
      } else {
        const sessionRes = await fetch(
          `https://graph.microsoft.com/v1.0/drives/${driveId}/root:/${filePath}:/createUploadSession`,
          {
            method: "POST",
            headers: { Authorization: `Bearer ${user.token}`, "Content-Type": "application/json" },
            body: JSON.stringify({ item: { "@microsoft.graph.conflictBehavior": "rename", name: fileName } }),
          }
        );
        const session = await sessionRes.json();
        if (!session.uploadUrl)
          throw new Error("Impossible de créer la session upload: " + JSON.stringify(session).slice(0, 200));
        const uploadUrl = session.uploadUrl;
        const ab = await file.arrayBuffer();
        let start = 0;
        let chunkNum = 0;
        const totalChunks = Math.ceil(file.size / CHUNK_SIZE);
        while (start < file.size) {
          const end = Math.min(start + CHUNK_SIZE, file.size);
          const chunk = ab.slice(start, end);
          chunkNum++;
          const chunkRes = await fetch(uploadUrl, {
            method: "PUT",
            headers: {
              "Content-Range": `bytes ${start}-${end - 1}/${file.size}`,
              "Content-Length": String(end - start),
              "Content-Type": file.type || "application/octet-stream",
            },
            body: chunk,
          });
          if (chunkRes.status === 200 || chunkRes.status === 201) {
            resultData = await chunkRes.json();
            break;
          } else if (chunkRes.status === 202) {
            start = end;
          } else {
            const errData = await chunkRes.json();
            throw new Error("Erreur chunk " + chunkNum + ": " + (errData.error?.message || chunkRes.status));
          }
        }
      }
      if (resultData?.id) {
        setDoc((d) => [
          ...d,
          {
            nom: resultData.name,
            url: resultData.webUrl,
            id: resultData.id,
            size: `${(file.size / 1024).toFixed(0)} Ko`,
          },
        ]);
        alert(`✅ "${file.name}" (${(file.size / 1024 / 1024).toFixed(1)} MB) uploadé dans SharePoint !`);
      } else {
        alert("Erreur: " + (resultData?.error?.message || JSON.stringify(resultData || {}).slice(0, 200)));
      }
    } catch (err) {
      alert("Erreur upload: " + err.message);
      console.error("Upload error:", err);
    }
    setUploading(false);
    e.target.value = "";
  };
  const loadDocs = async (token) => {
    try {
      const driveId = await getDriveId(token);
      if (!driveId) return;
      const folder = `ECOGEST_Etudiants/${etudiantFolderName(editItem)}`;
      const res = await fetch(`https://graph.microsoft.com/v1.0/drives/${driveId}/root:/${folder}:/children`, {
        headers: { Authorization: `Bearer ${token}` },
      });
      const data2 = await res.json();
      if (data2.value) setDoc(data2.value.map((f) => ({ nom: f.name, url: f.webUrl, id: f.id })));
    } catch (e) {
      console.warn("loadDocs error:", e.message);
    }
  };
  const handleDeleteDoc = async (att) => {
    if (!window.confirm('Supprimer "' + att.nom + '" ?')) return;
    try {
      await fetch(`https://graph.microsoft.com/v1.0/drives/${window._spDriveId}/items/${att.id}`, {
        method: "DELETE",
        headers: { Authorization: `Bearer ${user.token}` },
      });
      setDoc((d) => d.filter((x) => x.id !== att.id));
    } catch (e) {
      alert("Erreur suppression: " + e.message);
    }
  };
  const filtered = data.etudiants.filter(
    (e) =>
      `${e.nom} ${e.prenom} ${e.email}`.toLowerCase().includes(search.toLowerCase()) &&
      (!filterFiliere || e.filiere === filterFiliere) &&
      (!filterStatut || e.statut === filterStatut)
  );
  const filteredSorted = [...filtered].sort((a, b) => {
    const codeA = filiereCode(a.filiere, data.filieres) || a.filiere || "";
    const codeB = filiereCode(b.filiere, data.filieres) || b.filiere || "";
    return (
      codeA.localeCompare(codeB, "fr", { numeric: true, sensitivity: "base" }) ||
      (a.nom || "").localeCompare(b.nom || "", "fr", { sensitivity: "base" })
    );
  });
  const bandeFiliereParId = {};
  const premiereLigneFiliereId = {};
  const derniereLigneFiliereId = {};
  {
    let bande = 0,
      filierePrecedente = null;
    filteredSorted.forEach((e, idx) => {
      if (e.filiere !== filierePrecedente) {
        bande = 1 - bande;
        filierePrecedente = e.filiere;
        premiereLigneFiliereId[e.id] = true;
        if (idx > 0) derniereLigneFiliereId[filteredSorted[idx - 1].id] = true;
      }
      bandeFiliereParId[e.id] = bande;
    });
    if (filteredSorted.length > 0) derniereLigneFiliereId[filteredSorted[filteredSorted.length - 1].id] = true;
  }
  const handleImprimerListeFiliere = () => {
    if (!filterFiliere) {
      toast("⚠️ Sélectionnez un groupe dans le filtre pour imprimer sa liste");
      return;
    }
    const filiereObj = (data.filieres || []).find((f) => f.intitule === filterFiliere);
    const etudiantsGroupe = filteredSorted.filter((e) => e.filiere === filterFiliere);
    genererListeEtudiantsFiliere(filiereObj, etudiantsGroupe);
  };
  const openAdd = () => {
    const anneeActive = data.configuration?.anneeUniversitaireActive || ANNEES_UNIVERSITAIRES[0];
    setEditItem(null);
    setForm({
      ...emptyForm,
      numInscription: genNumInscriptionESRMI(data.etudiants, anneeActive, data.paiements?.etudiants_paiement),
      promotion: anneeActive,
    });
    setShowModal(true);
  };
  const openEdit = (et) => {
    setEditItem(et);
    setForm({ ...et });
    setShowModal(true);
  };
  const handleSave = () => {
    const spF = {
      Title: (form.nom + " " + form.prenom).toUpperCase(),
      Nom: (form.nom || "").toUpperCase(),
      Prenom: (form.prenom || "").toUpperCase(),
      CIN: (form.cin || "").toUpperCase(),
      DateNaissance: form.dateNaissance || "",
      LieuNaissance: form.lieuNaissance || "",
      Filiere: form.filiere || "",
      Niveau: form.niveau || "",
      Annee: form.annee || "",
      Email: form.email || "",
      EmailPersonnel: form.emailPersonnel || "",
      Telephone: form.telephone || "",
      Statut: form.statut || "Actif",
      Genre: form.genre || "M",
      Ville: form.ville || "",
      Quartier: form.quartier || "",
      Adresse: form.adresse || "",
      NumInscription: form.numInscription || "",
      Promotion: form.promotion || "",
      ResponsableFinancier: form.responsableFinancier || "",
      InscritPar: form.inscritPar || "",
      BacAnnee: form.bacAnnee || "",
      BacMention: form.bacMention || "",
      BacSerie: form.bacSerie || "",
      Etablissement: form.etablissement || "",
      Observations: form.observations || "",
      TarifGlobalFormation: String(form.tarifGlobalFormation || 0),
      ModeReglement: form.modeReglement || "Trimestriel",
      NotesConvention: form.notesConvention || "",
    };
    if (editItem) {
      let visiteurSync = null;
      setData((d) => {
        const linked = form.cin ? (d.visiteurs || []).find((v) => v.cin && v.cin === form.cin) : null;
        let visiteurs = d.visiteurs || [];
        if (linked) {
          visiteurSync = {
            ...linked,
            nom: form.nom || "",
            prenom: form.prenom || "",
            cin: form.cin || "",
            dateNaissance: form.dateNaissance || "",
            lieuNaissance: form.lieuNaissance || "",
            genre: form.genre || "M",
            telephone: form.telephone || "",
            email: form.email || "",
            adresse: form.adresse || "",
            ville: form.ville || "",
            quartier: form.quartier || "",
          };
          visiteurs = visiteurs.map((v) => (v.id === linked.id ? visiteurSync : v));
        }
        return {
          ...d,
          etudiants: d.etudiants.map((e) =>
            e.id === editItem.id ? { ...form, id: editItem.id, _spId: editItem._spId } : e
          ),
          visiteurs,
        };
      });
      if (user?.token && editItem._spId) saveToSP("ECOGEST_Etudiants", spF, user.token, editItem._spId);
      if (visiteurSync && user?.token && visiteurSync._spId) {
        saveToSP(
          "ECOGEST_Visiteurs",
          {
            Nom: visiteurSync.nom || "",
            Prenom: visiteurSync.prenom || "",
            CIN: visiteurSync.cin || "",
            DateNaissance: visiteurSync.dateNaissance || "",
            LieuNaissance: visiteurSync.lieuNaissance || "",
            Genre: visiteurSync.genre || "M",
            Telephone: visiteurSync.telephone || "",
            Email: visiteurSync.email || "",
            Adresse: visiteurSync.adresse || "",
            Ville: visiteurSync.ville || "",
            Quartier: visiteurSync.quartier || "",
          },
          user.token,
          visiteurSync._spId
        );
      }
      toast(
        visiteurSync
          ? "Étudiant et fiche visiteur liée mis à jour dans SharePoint ✓"
          : "Étudiant mis à jour dans SharePoint ✓"
      );
    } else {
      const newEt = { ...form, id: Date.now() };
      if (!user?.token) {
        toast("⚠️ Non connecté à SharePoint — étudiant non enregistré");
      } else {
        saveToSP("ECOGEST_Etudiants", spF, user.token).then((result) => {
          if (!result) {
            toast("⚠️ Échec de l'enregistrement dans SharePoint — étudiant non ajouté");
            return;
          }
          setData((d) => ({ ...d, etudiants: [...d.etudiants, { ...newEt, _spId: result.id }] }));
          toast("Étudiant enregistré dans SharePoint ✓");
        });
      }
    }
    setShowModal(false);
  };
  const handleDelete = (et) => {
    if (!window.confirm(`Supprimer l'étudiant "${et.prenom} ${et.nom}" ? Cette action est irréversible.`)) return;
    let visiteurRevert = null;
    setData((d) => {
      const linked = et.cin
        ? (d.visiteurs || []).find((v) => v.cin && v.cin === et.cin && v.statut === "Inscrit")
        : null;
      let visiteurs = d.visiteurs || [];
      if (linked) {
        visiteurRevert = { ...linked, statut: "Visiteur" };
        visiteurs = visiteurs.map((v) => (v.id === linked.id ? visiteurRevert : v));
      }
      return { ...d, etudiants: d.etudiants.filter((e) => e.id !== et.id), visiteurs };
    });
    if (user?.token && et._spId) deleteFromSP("ECOGEST_Etudiants", et._spId, user.token);
    if (visiteurRevert && user?.token && visiteurRevert._spId) {
      saveToSP("ECOGEST_Visiteurs", { Statut: "Visiteur" }, user.token, visiteurRevert._spId);
    }
    toast(
      visiteurRevert
        ? 'Étudiant supprimé — visiteur repassé en statut "Visiteur" ✓'
        : "Étudiant supprimé de SharePoint ✓"
    );
    setShowModal(false);
  };
  return (
    /* @__PURE__ */ <div>
      <PageHeader
        title="Étudiants"
        sub={`${filtered.length} étudiant(s) — données stockées dans SharePoint ECOGEST`}
        action={
          /* @__PURE__ */ <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
            <Btn variant="navy" onClick={handleImprimerListeFiliere}>
              🖨️ Imprimer la liste du groupe
            </Btn>
            {canEdit && /* @__PURE__ */ <Btn onClick={openAdd}>+ Ajouter un étudiant</Btn>}
          </div>
        }
      />
      <div style={{ display: "flex", gap: 10, marginBottom: 14, flexWrap: "wrap" }}>
        <input
          placeholder="🔍  Rechercher par nom ou email..."
          value={search}
          onChange={(e) => setSearch(e.target.value)}
          style={{
            flex: 1,
            minWidth: 220,
            padding: "10px 14px",
            border: `1px solid ${C.border}`,
            borderRadius: 8,
            fontSize: 13,
            boxSizing: "border-box",
            fontFamily: "inherit",
          }}
        />
        <select
          value={filterFiliere}
          onChange={(e) => setFilterFiliere(e.target.value)}
          style={{
            padding: "10px 14px",
            border: `1px solid ${C.border}`,
            borderRadius: 8,
            fontSize: 13,
            fontFamily: "inherit",
            background: C.white,
            minWidth: 260,
          }}
        >
          <option value="">-- Tous les groupes --</option>
          {getFilieres(data).map((f) => (
            /* @__PURE__ */ <option key={f} value={f}>
              {f}
            </option>
          ))}
        </select>
        <select
          value={filterStatut}
          onChange={(e) => setFilterStatut(e.target.value)}
          style={{
            padding: "10px 14px",
            border: `1px solid ${C.border}`,
            borderRadius: 8,
            fontSize: 13,
            fontFamily: "inherit",
            background: C.white,
          }}
        >
          <option value="">-- Tous les statuts --</option>
          <option value="Actif">Actif</option>
          <option value="Inactif">Inactif</option>
        </select>
      </div>
      {(() => {
        const cols = [
          {
            key: "_filiereCode",
            label: "Filière",
            render: (_, row) => (
              /* @__PURE__ */ <FiliereBadge code={filiereCode(row.filiere, data.filieres) || row.filiere} />
            ),
          },
          {
            key: "numInscription",
            label: "N° Inscription",
            render: (v, row) => (
              /* @__PURE__ */ <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
                <span>{v}</span>
                {(data.visiteurs || []).some((vi) => vi.cin && vi.cin === row.cin) && (
                  /* @__PURE__ */ <span
                    title="Étudiant converti depuis un visiteur"
                    style={{
                      background: C.purple,
                      color: "white",
                      fontSize: 9.5,
                      fontWeight: 800,
                      borderRadius: "50%",
                      width: 16,
                      height: 16,
                      display: "flex",
                      alignItems: "center",
                      justifyContent: "center",
                      flexShrink: 0,
                    }}
                  >
                    V
                  </span>
                )}
              </div>
            ),
          },
          { key: "prenom", label: "Prénom" },
          {
            key: "nom",
            label: "Nom",
            render: (v, row) => (
              /* @__PURE__ */ <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
                <span>{v}</span>
                {estEtudiantTest(row.prenom, row.nom) && (
                  /* @__PURE__ */ <span
                    title="Étudiant de test — n'existe pas réellement"
                    style={{
                      background: C.purple,
                      color: "white",
                      fontSize: 9.5,
                      fontWeight: 800,
                      borderRadius: 20,
                      padding: "1px 7px",
                      flexShrink: 0,
                      display: "inline-flex",
                      alignItems: "center",
                      gap: 3,
                    }}
                  >
                    🧪 TEST
                  </span>
                )}
              </div>
            ),
          },
          {
            key: "cin",
            label: "CIN",
            render: (v) => (
              /* @__PURE__ */ <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
                <span>{v || "—"}</span>
                {(!v || v.trim() === "" || v.trim() === "-") && (
                  /* @__PURE__ */ <span
                    title="CIN non défini"
                    style={{
                      fontSize: 14,
                      cursor: "help",
                      display: "inline-block",
                      animation: "pulseWarn 1.2s ease-in-out infinite",
                    }}
                  >
                    ⚠️
                  </span>
                )}
              </div>
            ),
          },
          { key: "email", label: "Email" },
          {
            key: "tarifGlobalFormation",
            label: "Tarif de formation",
            render: (v, row) => {
              const coul = colorForFiliereCode(filiereCode(row.filiere, data.filieres) || row.filiere);
              return (
                /* @__PURE__ */ <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
                  <span
                    style={{
                      display: "inline-flex",
                      alignItems: "center",
                      gap: 6,
                      background: v ? coul : C.muted,
                      color: "#ffffff",
                      fontSize: 11,
                      fontWeight: 700,
                      padding: "3px 10px 3px 4px",
                      borderRadius: 20,
                      flexShrink: 0,
                      whiteSpace: "nowrap",
                      boxShadow: v ? `0 3px 8px ${coul}80` : "none",
                    }}
                  >
                    <span
                      style={{
                        background: "rgba(255,255,255,0.94)",
                        color: v ? coul : C.muted,
                        borderRadius: "50%",
                        width: 19,
                        height: 19,
                        display: "flex",
                        alignItems: "center",
                        justifyContent: "center",
                        fontSize: 11,
                        flexShrink: 0,
                      }}
                    >
                      💰
                    </span>
                    {v ? formatMAD(v) : "Non défini"}
                  </span>
                  {!v && (
                    /* @__PURE__ */ <span
                      title="Tarif de la formation non défini dans la Convention financière"
                      style={{
                        fontSize: 14,
                        cursor: "help",
                        display: "inline-block",
                        animation: "pulseWarn 1.2s ease-in-out infinite",
                      }}
                    >
                      ⚠️
                    </span>
                  )}
                </div>
              );
            },
          },
          {
            key: "_docs",
            label: "Documents",
            headerAlign: "center",
            render: (_, row) => (
              /* @__PURE__ */ <div style={{ display: "flex", gap: 6 }} onClick={(e) => e.stopPropagation()}>
                <DocButton
                  icon="📄"
                  label="Inscription"
                  onClick={() => genererAttestation(row, row.filiere, "inscription")}
                />
                <DocButton
                  icon="📜"
                  label="Scolarité"
                  onClick={() => genererAttestation(row, row.filiere, "scolarite")}
                />
                <DocButton
                  icon="🏆"
                  label="Réussite"
                  onClick={() => genererAttestation(row, row.filiere, "reussite")}
                />
                <DocButton
                  icon="📁"
                  label="Pièces"
                  onClick={() => openEtudiantFolder(row)}
                  badge={docCounts[row.id] > 0 ? docCounts[row.id] : null}
                />
              </div>
            ),
          },
          ...(canEdit
            ? [
                {
                  key: "_modifier",
                  label: "Modifier",
                  render: (_, row) => (
                    /* @__PURE__ */ <Btn
                      small={true}
                      variant="navy"
                      onClick={(e) => {
                        e.stopPropagation();
                        openEdit(row);
                      }}
                    >
                      ✏️ Modifier
                    </Btn>
                  ),
                },
              ]
            : []),
          ...(canEdit
            ? [
                {
                  key: "id",
                  label: "Suppr.",
                  render: (_, row) => (
                    /* @__PURE__ */ <Btn
                      small={true}
                      variant="danger"
                      onClick={(e) => {
                        e.stopPropagation();
                        handleDelete(row);
                      }}
                    >
                      🗑
                    </Btn>
                  ),
                },
              ]
            : []),
        ];
        if (filteredSorted.length === 0)
          return (
            /* @__PURE__ */ <Card style={{ padding: 32, textAlign: "center", color: C.muted, fontSize: 13 }}>
              Aucun étudiant trouvé
            </Card>
          );
        return (
          /* @__PURE__ */ <Card style={{ overflow: "hidden" }}>
            <DataTable
              cols={cols}
              rows={filteredSorted}
              compact={true}
              emptyMsg="Aucun étudiant trouvé"
              rowStyle={(row) => ({
                background: bandeFiliereParId[row.id] === 0 ? C.navyLight : "#e0f4f2",
                borderTop: premiereLigneFiliereId[row.id] ? `2.5px solid ${C.navy}` : void 0,
                borderBottom: derniereLigneFiliereId[row.id] ? `2.5px solid ${C.navy}` : `1px solid ${C.border}`,
              })}
            />
          </Card>
        );
      })()}
      {showModal && (
        /* @__PURE__ */ <Modal
          title={editItem ? "Modifier l'étudiant → SharePoint" : "Ajouter un étudiant → SharePoint"}
          onClose={() => setShowModal(false)}
          hideFooterClose={true}
        >
          <div className="eco-grid2" style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "0 14px" }}>
            <Input label="Nom *" value={form.nom} onChange={(e) => setForm((f) => ({ ...f, nom: e.target.value }))} />
            <Input
              label="Prénom *"
              value={form.prenom}
              onChange={(e) => setForm((f) => ({ ...f, prenom: e.target.value }))}
            />
          </div>
          <div className="eco-grid2" style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "0 14px" }}>
            <Input
              label="Email scolaire *"
              type="email"
              value={form.email}
              onChange={(e) => setForm((f) => ({ ...f, email: e.target.value }))}
              placeholder="etudiant@esrmi.ma"
            />
            <Input
              label="Email personnel"
              type="email"
              value={form.emailPersonnel}
              onChange={(e) => setForm((f) => ({ ...f, emailPersonnel: e.target.value }))}
              placeholder="etudiant@gmail.com"
            />
          </div>
          <div>
            <Input
              label="Téléphone *"
              value={form.telephone}
              onChange={(e) => setForm((f) => ({ ...f, telephone: e.target.value }))}
              placeholder="06XXXXXXXX"
              style={form.telephone && !/^\d{10}$/.test(form.telephone) ? { borderColor: C.red } : {}}
            />
            {form.telephone && !/^\d{10}$/.test(form.telephone) && (
              /* @__PURE__ */ <div style={{ fontSize: 11, color: C.red, marginTop: -10, marginBottom: 13 }}>
                ⚠️ Le numéro doit contenir exactement 10 chiffres
              </div>
            )}
          </div>
          <div className="eco-grid2" style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "0 14px" }}>
            <Input
              label="N° Inscription (généré automatiquement)"
              value={form.numInscription}
              readOnly={true}
              disabled={true}
              style={{ background: C.light, color: C.muted, cursor: "not-allowed" }}
            />
            <Input
              label="CIN *"
              value={form.cin}
              onChange={(e) => setForm((f) => ({ ...f, cin: e.target.value }))}
              placeholder="BK123456"
            />
          </div>
          <div className="eco-grid2" style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "0 14px" }}>
            <Input
              label="Date de naissance"
              type="date"
              value={form.dateNaissance}
              onChange={(e) => setForm((f) => ({ ...f, dateNaissance: e.target.value }))}
            />
            <Input
              label="Lieu de naissance"
              value={form.lieuNaissance}
              onChange={(e) => setForm((f) => ({ ...f, lieuNaissance: e.target.value }))}
            />
          </div>
          <div className="eco-grid3" style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: "0 14px" }}>
            <Input
              label="Adresse"
              value={form.adresse}
              onChange={(e) => setForm((f) => ({ ...f, adresse: e.target.value }))}
            />
            <Input
              label="Ville"
              value={form.ville}
              onChange={(e) => setForm((f) => ({ ...f, ville: e.target.value }))}
            />
            <Input
              label="Quartier"
              value={form.quartier}
              onChange={(e) => setForm((f) => ({ ...f, quartier: e.target.value }))}
            />
          </div>
          <div
            style={{
              fontSize: 11,
              fontWeight: 700,
              color: C.orange,
              textTransform: "uppercase",
              letterSpacing: "0.08em",
              margin: "10px 0",
            }}
          >
            Baccalauréat & Parcours
          </div>
          <div className="eco-grid3" style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: "0 14px" }}>
            <Input
              label="Année du Bac"
              value={form.bacAnnee}
              onChange={(e) => setForm((f) => ({ ...f, bacAnnee: e.target.value }))}
              placeholder="2024"
            />
            <Sel
              label="Mention"
              value={form.bacMention}
              onChange={(e) => setForm((f) => ({ ...f, bacMention: e.target.value }))}
            >
              <option>Très bien</option>
              <option>Bien</option>
              <option>Assez bien</option>
              <option>Passable</option>
            </Sel>
            <Input
              label="Série du Bac"
              value={form.bacSerie}
              onChange={(e) => setForm((f) => ({ ...f, bacSerie: e.target.value }))}
              placeholder="Sciences Éco..."
            />
          </div>
          <Input
            label="Établissement d'origine"
            value={form.etablissement}
            onChange={(e) => setForm((f) => ({ ...f, etablissement: e.target.value }))}
          />
          <div className="eco-grid2" style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "0 14px" }}>
            <Sel
              label="Filière"
              value={form.filiere}
              onChange={(e) => setForm((f) => ({ ...f, filiere: e.target.value }))}
            >
              <option value="">-- Sélectionner --</option>
              {getFilieres(data).map((f) => (
                /* @__PURE__ */ <option key={f} value={f}>
                  {f}
                </option>
              ))}
            </Sel>
            <Sel label="Genre" value={form.genre} onChange={(e) => setForm((f) => ({ ...f, genre: e.target.value }))}>
              <option value="M">Homme</option>
              <option value="F">Femme</option>
            </Sel>
          </div>
          <Sel label="Statut" value={form.statut} onChange={(e) => setForm((f) => ({ ...f, statut: e.target.value }))}>
            <option>Actif</option>
            <option>Inactif</option>
            <option>Suspendu</option>
          </Sel>
          <div
            style={{
              fontSize: 11,
              fontWeight: 700,
              color: C.orange,
              textTransform: "uppercase",
              letterSpacing: "0.08em",
              margin: "10px 0",
            }}
          >
            Convention financière
          </div>
          <Input
            label="Tarif global de la formation (MAD)"
            type="number"
            min="0"
            value={form.tarifGlobalFormation || 0}
            onChange={(e) => setForm((f) => ({ ...f, tarifGlobalFormation: parseFloat(e.target.value) || 0 }))}
          />
          <Sel
            label="Mode de règlement"
            value={form.modeReglement}
            onChange={(e) => setForm((f) => ({ ...f, modeReglement: e.target.value }))}
          >
            <option>Annuel</option>
            <option>Semestriel</option>
            <option>Trimestriel</option>
            <option>Mensuel</option>
          </Sel>
          <Input
            label="Responsable financier (Nom et prénom)"
            value={form.responsableFinancier}
            onChange={(e) => setForm((f) => ({ ...f, responsableFinancier: e.target.value }))}
            placeholder="Ex: Hassan Benali (Père)"
          />
          <Sel
            label="Inscrit par"
            value={form.inscritPar}
            onChange={(e) => setForm((f) => ({ ...f, inscritPar: e.target.value }))}
          >
            <option value="">-- Sélectionner --</option>
            <option>ADILA BAHALLA</option>
            <option>KAWTAR BOUABID</option>
            <option>FATIMA ZAHRA ZNIBER</option>
          </Sel>
          <Input
            label="Notes convention"
            value={form.notesConvention}
            onChange={(e) => setForm((f) => ({ ...f, notesConvention: e.target.value }))}
            placeholder="Ex: Convention signée en présence des parents..."
          />
          <Input
            label="Observations"
            value={form.observations}
            onChange={(e) => setForm((f) => ({ ...f, observations: e.target.value }))}
          />
          <div
            style={{
              background: "#f0f7ff",
              border: "1px solid #bee3f8",
              borderRadius: 8,
              padding: "10px 14px",
              fontSize: 12,
              color: "#2b6cb0",
              marginBottom: 16,
            }}
          >
            💰 Total convention : <strong>{(form.tarifGlobalFormation || 0).toLocaleString()} MAD</strong> · Mode : <strong>{form.modeReglement}</strong>
          </div>
          {editItem && (
            /* @__PURE__ */ <div style={{ marginBottom: 16 }}>
              <div
                style={{
                  fontSize: 11,
                  fontWeight: 700,
                  color: C.orange,
                  textTransform: "uppercase",
                  letterSpacing: "0.08em",
                  marginBottom: 10,
                }}
              >
                Pièces jointes SharePoint {loadingDocs && "⏳"}
              </div>
              <div
                style={{
                  background: "#f0f7ff",
                  border: "1px solid #bee3f8",
                  borderRadius: 8,
                  padding: "10px 14px",
                  fontSize: 12,
                  color: "#2b6cb0",
                  marginBottom: 10,
                }}
              >
                📁 Les documents sont stockés dans <strong>SharePoint → Documents partagés → ECOGEST_Etudiants → {etudiantFolderName(editItem)}</strong>
                <br />
                <a
                  href="https://ecolesuperieurederabat.sharepoint.com/sites/ECOGEST/Documents%20partages/Forms/AllItems.aspx"
                  target="_blank"
                  rel="noreferrer"
                  style={{ color: C.navy, fontWeight: 600, textDecoration: "none" }}
                >
                  📂 Ouvrir le dossier SharePoint
                </a>
              </div>
              <div
                className="eco-grid2"
                style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8, marginBottom: 10 }}
              >
                {[
                  "CIN / Carte d'identité",
                  "Relevé de notes Bac",
                  "Diplôme ou attestation",
                  "Photo d'identité",
                  "Autre document",
                ].map((type) => (
                  /* @__PURE__ */ <label
                    key={type}
                    style={{
                      display: "flex",
                      alignItems: "center",
                      gap: 8,
                      padding: "8px 12px",
                      border: `1px dashed ${C.border}`,
                      borderRadius: 8,
                      cursor: "pointer",
                      fontSize: 12,
                      color: C.navy,
                      background: C.light,
                    }}
                  >
                    <span>📎</span> {uploading ? "⏳ Upload en cours..." : type}
                    <input
                      type="file"
                      style={{ display: "none" }}
                      disabled={uploading}
                      onChange={(e) => handleUpload(e, type)}
                      accept=".pdf,.jpg,.jpeg,.png,.doc,.docx,.zip,.rar,.mp4,.avi"
                    />
                  </label>
                ))}
              </div>
              {doc.length > 0 ? (
                /* @__PURE__ */ <div style={{ border: `1px solid ${C.border}`, borderRadius: 8, overflow: "hidden" }}>
                  {doc.map((d, i) => (
                    /* @__PURE__ */ <div
                      key={i}
                      style={{
                        display: "flex",
                        justifyContent: "space-between",
                        alignItems: "center",
                        padding: "8px 12px",
                        background: i % 2 === 0 ? C.white : C.light,
                      }}
                    >
                      <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                        <span style={{ fontSize: 18 }}>📄</span>
                        <div style={{ fontSize: 12, fontWeight: 600, color: C.navy }}>{d.nom}</div>
                      </div>
                      <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
                        {d.url && (
                          /* @__PURE__ */ <a
                            href={d.url}
                            target="_blank"
                            rel="noreferrer"
                            style={{
                              fontSize: 12,
                              color: C.navy,
                              textDecoration: "none",
                              padding: "4px 10px",
                              border: `1px solid ${C.border}`,
                              borderRadius: 6,
                            }}
                          >
                            👁 Voir
                          </a>
                        )}
                        {canEdit && (
                          /* @__PURE__ */ <button
                            onClick={() => handleDeleteDoc(d)}
                            style={{
                              fontSize: 12,
                              color: C.red,
                              background: "none",
                              border: "1px solid #fecaca",
                              borderRadius: 6,
                              padding: "4px 10px",
                              cursor: "pointer",
                            }}
                          >
                            🗑
                          </button>
                        )}
                      </div>
                    </div>
                  ))}
                </div>
              ) : (
                /* @__PURE__ */ <div
                  style={{
                    fontSize: 12,
                    color: C.muted,
                    textAlign: "center",
                    padding: 16,
                    border: `1px dashed ${C.border}`,
                    borderRadius: 8,
                  }}
                >
                  {loadingDocs ? "Chargement des documents..." : "Aucun document uploadé dans SharePoint"}
                </div>
              )}
            </div>
          )}
          <div style={{ display: "flex", gap: 10, justifyContent: "space-between" }}>
            {editItem ? (
              /* @__PURE__ */ <Btn variant="danger" onClick={() => handleDelete(editItem)}>
                🗑 Supprimer
              </Btn>
            ) : (
              /* @__PURE__ */ <span />
            )}
            <div style={{ display: "flex", gap: 10 }}>
              <Btn variant="light" onClick={() => setShowModal(false)}>
                Annuler
              </Btn>
              <Btn
                onClick={handleSave}
                disabled={!form.nom || !form.prenom || !form.cin || !/^\d{10}$/.test(form.telephone || "")}
              >
                Enregistrer → SharePoint
              </Btn>
            </div>
          </div>
        </Modal>
      )}
    </div>
  );
}
function Professeurs({ data, setData, role, user, toast }) {
  const canEdit = ["administrateur", "directrice"].includes(role);
  const [showModal, setShowModal] = useState(false);
  const [editItem, setEditItem] = useState(null);
  const [search, setSearch] = useState("");
  const emptyForm = { nom: "", prenom: "", cin: "", email: "", telephone: "", statut: "Actif", type: "Permanent" };
  const [form, setForm] = useState(emptyForm);
  const extractCinFromProfesseurStr = (s) => {
    const m = /\(CIN:\s*([^)]+)\)/.exec(s || "");
    return m ? m[1].trim().toUpperCase() : "";
  };
  const matieresDe = (p) => [
    ...new Set(
      (data.cours || [])
        .filter(
          (c) =>
            extractCinFromProfesseurStr(c.professeur) &&
            extractCinFromProfesseurStr(c.professeur) === (p.cin || "").trim().toUpperCase()
        )
        .map((c) => c.intitule)
        .filter(Boolean)
    ),
  ];
  const openAdd = () => {
    setEditItem(null);
    setForm(emptyForm);
    setShowModal(true);
  };
  const openEdit = (p) => {
    setEditItem(p);
    setForm({ ...emptyForm, ...p });
    setShowModal(true);
  };
  const spFields = (f) => ({
    Title: (f.nom || "") + " " + (f.prenom || ""),
    Nom: f.nom || "",
    Prenom: f.prenom || "",
    CIN: f.cin || "",
    Email: f.email || "",
    Telephone: f.telephone || "",
    TypeContrat: f.type || "Permanent",
    Statut: f.statut || "Actif",
  });
  const handleSave = () => {
    if (!user?.token) {
      toast("⚠️ Non connecté à SharePoint — professeur non enregistré");
      setShowModal(false);
      return;
    }
    const cinSaisi = (form.cin || "").trim().toUpperCase();
    const cinDejaUtilise = (data.professeurs || []).some(
      (p) => p.id !== editItem?.id && (p.cin || "").trim().toUpperCase() === cinSaisi
    );
    if (cinDejaUtilise) {
      toast(`⚠️ Un professeur avec le CIN "${form.cin}" existe déjà — le CIN doit être unique`);
      return;
    }
    if (editItem) {
      const updated = { ...form, id: editItem.id, _spId: editItem._spId };
      setData((d) => ({ ...d, professeurs: d.professeurs.map((p) => (p.id === editItem.id ? updated : p)) }));
      saveToSP("ECOGEST_Professeurs", spFields(form), user.token, editItem._spId).then((result) => {
        if (!result) {
          toast("⚠️ Échec de la mise à jour dans SharePoint");
          return;
        }
        toast("Professeur mis à jour dans SharePoint ✓");
      });
    } else {
      const newP = { ...form, id: Date.now() };
      saveToSP("ECOGEST_Professeurs", spFields(form), user.token).then((result) => {
        if (!result) {
          toast("⚠️ Échec de l'enregistrement dans SharePoint — professeur non ajouté");
          return;
        }
        setData((d) => ({ ...d, professeurs: [...d.professeurs, { ...newP, _spId: result.id }] }));
        toast("Professeur enregistré dans SharePoint ✓");
      });
    }
    setShowModal(false);
    setForm(emptyForm);
  };
  const handleDelete = (p) => {
    if (!window.confirm(`Supprimer le professeur "${p.prenom} ${p.nom}" ?`)) return;
    setData((d) => ({ ...d, professeurs: d.professeurs.filter((x) => x.id !== p.id) }));
    if (user?.token && p._spId) deleteFromSP("ECOGEST_Professeurs", p._spId, user.token);
    toast("Professeur supprimé de SharePoint ✓");
  };
  const cols = [
    { key: "cin", label: "CIN" },
    { key: "prenom", label: "Prénom" },
    { key: "nom", label: "Nom" },
    {
      key: "matiere",
      label: "Matière(s) enseignée(s)",
      render: (_, row) => {
        const m = matieresDe(row);
        return m.length ? m.join(", ") : /* @__PURE__ */ <span style={{ color: C.muted, fontStyle: "italic" }}>—</span>;
      },
    },
    {
      key: "type",
      label: "Type",
      render: (v) => /* @__PURE__ */ <Badge color={v === "Permanent" ? C.navy : C.orange}>{v}</Badge>,
    },
    { key: "email", label: "Email" },
    {
      key: "statut",
      label: "Statut",
      render: (v) => /* @__PURE__ */ <Badge color={v === "Actif" ? C.green : C.muted}>{v}</Badge>,
    },
    ...(canEdit
      ? [
          {
            key: "_modifier",
            label: "Modifier",
            render: (_, row) => (
              /* @__PURE__ */ <Btn small={true} variant="navy" onClick={() => openEdit(row)}>
                ✏️ Modifier
              </Btn>
            ),
          },
        ]
      : []),
    ...(canEdit
      ? [
          {
            key: "_suppr",
            label: "Suppr.",
            render: (_, row) => (
              /* @__PURE__ */ <Btn small={true} variant="danger" onClick={() => handleDelete(row)}>
                🗑
              </Btn>
            ),
          },
        ]
      : []),
  ];
  const professeursFiltres = (data.professeurs || [])
    .filter(
      (p) =>
        !search ||
        `${p.nom || ""} ${p.prenom || ""} ${p.cin || ""} ${p.email || ""}`.toLowerCase().includes(search.toLowerCase())
    )
    .sort(
      (a, b) =>
        (a.prenom || "").localeCompare(b.prenom || "", "fr", { sensitivity: "base" }) ||
        (a.nom || "").localeCompare(b.nom || "", "fr", { sensitivity: "base" })
    );
  return (
    /* @__PURE__ */ <div>
      <PageHeader
        title="Corps enseignant"
        sub={`${professeursFiltres.length} professeur(s) — liste SharePoint "${SP.lists.professeurs}"`}
        action={canEdit && /* @__PURE__ */ <Btn onClick={openAdd}>+ Ajouter un professeur</Btn>}
      />
      <div style={{ display: "flex", gap: 10, marginBottom: 14 }}>
        <input
          placeholder="🔍  Nom, prénom, CIN, email..."
          value={search}
          onChange={(e) => setSearch(e.target.value)}
          style={{
            flex: 1,
            padding: "10px 14px",
            border: `1px solid ${C.border}`,
            borderRadius: 8,
            fontSize: 13,
            boxSizing: "border-box",
            fontFamily: "inherit",
          }}
        />
      </div>
      <Card style={{ overflow: "hidden" }}>
        <DataTable cols={cols} rows={professeursFiltres} />
      </Card>
      {showModal && (
        /* @__PURE__ */ <Modal
          title={editItem ? "Modifier le professeur" : "Ajouter un professeur → SharePoint"}
          onClose={() => setShowModal(false)}
        >
          <div className="eco-grid2" style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "0 14px" }}>
            <Input label="Nom *" value={form.nom} onChange={(e) => setForm((f) => ({ ...f, nom: e.target.value }))} />
            <Input
              label="Prénom *"
              value={form.prenom}
              onChange={(e) => setForm((f) => ({ ...f, prenom: e.target.value }))}
            />
          </div>
          <div className="eco-grid2" style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "0 14px" }}>
            <Input label="CIN *" value={form.cin} onChange={(e) => setForm((f) => ({ ...f, cin: e.target.value }))} />
            <Input
              label="Email *"
              value={form.email}
              onChange={(e) => setForm((f) => ({ ...f, email: e.target.value }))}
            />
          </div>
          <Input
            label="Téléphone"
            value={form.telephone}
            onChange={(e) => setForm((f) => ({ ...f, telephone: e.target.value }))}
          />
          {editItem && (
            /* @__PURE__ */ <div
              style={{
                background: C.light,
                borderRadius: 8,
                padding: "8px 12px",
                fontSize: 11.5,
                color: "#666",
                marginBottom: 14,
              }}
            >
              ℹ️ La matière enseignée est calculée automatiquement à partir des modules assignés à ce professeur (page
              Modules).
            </div>
          )}
          <div className="eco-grid2" style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "0 14px" }}>
            <Sel label="Type" value={form.type} onChange={(e) => setForm((f) => ({ ...f, type: e.target.value }))}>
              <option>Permanent</option>
              <option>Vacataire</option>
            </Sel>
            <Sel
              label="Statut"
              value={form.statut}
              onChange={(e) => setForm((f) => ({ ...f, statut: e.target.value }))}
            >
              <option>Actif</option>
              <option>Inactif</option>
            </Sel>
          </div>
          <div style={{ display: "flex", gap: 10, justifyContent: "flex-end" }}>
            <Btn variant="light" onClick={() => setShowModal(false)}>
              Annuler
            </Btn>
            <Btn onClick={handleSave} disabled={!form.nom || !form.prenom || !form.cin}>
              {editItem ? "Mettre à jour → SharePoint" : "Enregistrer → SharePoint"}
            </Btn>
          </div>
        </Modal>
      )}
    </div>
  );
}
const TYPES_EVALUATION = ["Contrôle continu", "Examen de fin de module", "Rattrapage"];
function Notes({ data, setData, role, user, toast }) {
  const canEdit = ["administrateur", "professeur"].includes(role);
  const canPrint = ["administrateur", "directrice", "professeur"].includes(role);
  const [showModal, setShowModal] = useState(false);
  const [saving, setSaving] = useState(false);
  const emptyForm = { date: today(), semestre: "", filiere: "", module: "", typeEvaluation: "Contrôle continu" };
  const [form, setForm] = useState(emptyForm);
  const [notesSaisies, setNotesSaisies] = useState({});
  const [notesIdsEdition, setNotesIdsEdition] = useState({});
  const [vueOuverte, setVueOuverte] = useState({});
  const monEtudiant = (data.etudiants || []).find(
    (e) =>
      (user?.numInscription && e.numInscription === user.numInscription) ||
      (user?.etudiantId && e.etudiantId === user.etudiantId)
  );
  const monProfesseur = (data.professeurs || []).find(
    (p) => p.email && user?.email && p.email.toLowerCase() === user.email.toLowerCase()
  );
  const mesNotes =
    role === "etudiant"
      ? (data.notes || []).filter(
          (n) =>
            (user?.numInscription && n.numInscription === user.numInscription) ||
            (user?.etudiantId && n.etudiantId === user.etudiantId)
        )
      : role === "professeur"
        ? (data.notes || []).filter((n) => n.enseignantCin === monProfesseur?.cin)
        : data.notes || [];
  const filieresOptions = getFilieres(data);
  const semestresOptions = [
    .../* @__PURE__ */ new Set([
      ...SEMESTRES_MODULE_SEED,
      ...(data.cours || []).map((c) => c.semestre).filter(Boolean),
    ]),
  ].sort((a, b) => a.localeCompare(b, "fr", { numeric: true, sensitivity: "base" }));
  const modulesOptions = (data.cours || []).filter(
    (c) =>
      (role === "professeur" ? c.professeurCin === monProfesseur?.cin : true) &&
      (!form.filiere || (c.filieresListe || [c.filiere]).includes(form.filiere)) &&
      (!form.semestre || c.semestre === form.semestre)
  );
  const moduleChoisi = (data.cours || []).find((c) => c.intitule === form.module);
  const semestreInfoForm = (data.semestresDates || []).find((s) => s.semestre === form.semestre);
  const etudiantsFiliere = form.filiere ? (data.etudiants || []).filter((e) => e.filiere === form.filiere) : [];
  const imprimerRelevesGroupe = (type) => {
    const paires = [...new Map(lignesEvaluations.map((n) => [n.numInscription + "|||" + n.semestre, n])).values()];
    if (paires.length === 0) {
      toast("⚠️ Aucun étudiant ne correspond aux filtres actuels");
      return;
    }
    const W = 210 * MM,
      H = 297 * MM;
    const doc = new MiniPDF(W, H);
    paires.forEach((row, i) => {
      if (i > 0) doc.newPage();
      const et = (data.etudiants || []).find((e) => e.numInscription === row.numInscription) || {
        prenom: row.etudiant || "",
        nom: "",
        filiere: row._filiere || "",
        niveau: "",
      };
      const semestre = row.semestre || "";
      const notesSem = (data.notes || []).filter(
        (n) => n.numInscription === row.numInscription && n.semestre === semestre
      );
      const anneeUniv = data.configuration?.anneeUniversitaireActive || ANNEES_UNIVERSITAIRES[0];
      if (type === "synthetique") {
        const presencesSem = (data.presences || []).filter((p) => p.numInscription === row.numInscription);
        genererBulletinNotes(
          et,
          notesSem,
          semestre,
          anneeUniv,
          data.cours,
          presencesSem,
          data.configuration?.filieresAvecTauxAbsence,
          doc
        );
      } else {
        genererReleveDetaille(et, notesSem, semestre, anneeUniv, doc);
      }
    });
    const suffixe = type === "synthetique" ? "Synthetiques" : "Detailles";
    downloadPDF(doc.build(), `Releves_${suffixe}_${paires.length}_etudiants_${today()}`.replace(/\s+/g, "_") + ".pdf");
    toast(`✅ PDF généré — ${paires.length} relevé(s) dans un seul fichier`);
  };
  const openAdd = () => {
    setForm(emptyForm);
    setNotesSaisies({});
    setNotesIdsEdition({});
    setShowModal(true);
  };
  const openEditSession = (dateVal, filiereVal, semestreVal, moduleVal, typeVal, entries) => {
    setForm({ date: dateVal, semestre: semestreVal, filiere: filiereVal, module: moduleVal, typeEvaluation: typeVal });
    const saisies = {},
      ids = {};
    entries.forEach((n) => {
      saisies[n.etudiantId] = String(n.note);
      ids[n.etudiantId] = { id: n.id, _spId: n._spId };
    });
    setNotesSaisies(saisies);
    setNotesIdsEdition(ids);
    setShowModal(true);
  };
  const nbNotesSaisies = Object.values(notesSaisies).filter((v) => v !== "" && v !== void 0).length;
  const handleSaveNotes = async () => {
    if (!user?.token) {
      toast("⚠️ Non connecté à SharePoint — notes non enregistrées");
      return;
    }
    const toutes = Object.entries(notesSaisies).filter(([, v]) => v !== "" && v !== void 0 && v !== null);
    const entrees = toutes.filter(([, v]) => !isNaN(parseFloat(v)) && parseFloat(v) >= 0 && parseFloat(v) <= 20);
    if (toutes.length === 0) {
      toast("⚠️ Saisissez au moins une note");
      return;
    }
    if (entrees.length === 0) {
      toast("⚠️ Toutes les notes saisies sont hors de la plage 0-20 — corrigez-les");
      return;
    }
    if (entrees.length < toutes.length)
      toast(`⚠️ ${toutes.length - entrees.length} note(s) hors plage 0-20 ignorée(s)`);
    setSaving(true);
    const enregistrees = [];
    const misesAJour = [];
    for (const [etId, valeur] of entrees) {
      const et = etudiantsFiliere.find((e) => String(e.id) === etId);
      if (!et) continue;
      const spF = {
        Title: `${et.nom} ${et.prenom} - ${form.module} - ${form.typeEvaluation}`,
        EtudiantId: String(et.id),
        NomEtudiant: `${et.prenom} ${et.nom}`,
        NumInscription: et.numInscription || "",
        Module: form.module || "",
        Filiere: form.filiere || "",
        Semestre: form.semestre || "",
        TypeEvaluation: form.typeEvaluation || "",
        Note: String(parseFloat(valeur)),
        DateEvaluation: form.date || "",
        Enseignant: moduleChoisi?.professeur || "",
        EnseignantCIN: moduleChoisi?.professeurCin || "",
      };
      const existant = notesIdsEdition[etId];
      if (existant) {
        const result = await saveToSP("ECOGEST_Notes", spF, user.token, existant._spId);
        if (result) misesAJour.push({ id: existant.id, note: parseFloat(valeur), date: form.date });
      } else {
        const result = await saveToSP("ECOGEST_Notes", spF, user.token);
        if (result)
          enregistrees.push({
            id: Date.now() + Math.random(),
            _spId: result.id,
            etudiantId: et.id,
            etudiant: `${et.prenom} ${et.nom}`,
            numInscription: et.numInscription,
            module: form.module,
            filiere: form.filiere,
            semestre: form.semestre,
            typeEvaluation: form.typeEvaluation,
            note: parseFloat(valeur),
            date: form.date,
            enseignant: moduleChoisi?.professeur || "",
            enseignantCin: moduleChoisi?.professeurCin || "",
          });
      }
    }
    setSaving(false);
    if (enregistrees.length > 0) setData((d) => ({ ...d, notes: [...d.notes, ...enregistrees] }));
    if (misesAJour.length > 0)
      setData((d) => ({
        ...d,
        notes: d.notes.map((n) => {
          const maj = misesAJour.find((m) => m.id === n.id);
          return maj ? { ...n, note: maj.note, date: maj.date } : n;
        }),
      }));
    if (enregistrees.length + misesAJour.length < entrees.length)
      toast(
        `⚠️ ${entrees.length - enregistrees.length - misesAJour.length} note(s) non enregistrée(s) — échec SharePoint`
      );
    else toast(`${enregistrees.length + misesAJour.length} note(s) enregistrée(s) dans SharePoint ✓`);
    if (enregistrees.length + misesAJour.length > 0) {
      setShowModal(false);
      setForm(emptyForm);
      setNotesSaisies({});
      setNotesIdsEdition({});
    }
  };
  const handleDeleteNote = (n) => {
    if (!window.confirm(`Supprimer la note de "${n.etudiant}" (${n.module}) ?`)) return;
    setData((d) => ({ ...d, notes: d.notes.filter((x) => x.id !== n.id) }));
    if (user?.token && n._spId) deleteFromSP("ECOGEST_Notes", n._spId, user.token);
    toast("Note supprimée de SharePoint ✓");
  };
  const [filtreFiliereListe, setFiltreFiliereListe] = useState("");
  const [filtreSemestreListe, setFiltreSemestreListe] = useState("");
  const [filtreModuleListe, setFiltreModuleListe] = useState("");
  const [filtreTypeListe, setFiltreTypeListe] = useState("");
  const [searchListe, setSearchListe] = useState("");
  const modulesOptionsListe = [...new Set((data.notes || []).map((n) => n.module).filter(Boolean))].sort((a, b) =>
    a.localeCompare(b, "fr", { sensitivity: "base" })
  );
  const notesAffichees =
    role === "etudiant"
      ? mesNotes
      : mesNotes.filter(
          (n) =>
            (!filtreFiliereListe || n.filiere === filtreFiliereListe) &&
            (!filtreSemestreListe || n.semestre === filtreSemestreListe) &&
            (!filtreModuleListe || n.module === filtreModuleListe) &&
            (!filtreTypeListe || n.typeEvaluation === filtreTypeListe) &&
            (!searchListe ||
              `${n.etudiant || ""} ${n.numInscription || ""}`.toLowerCase().includes(searchListe.toLowerCase()))
        );
  const codeFiliere = (intitule) => (data.filieres || []).find((x) => x.intitule === intitule)?.code || intitule || "—";
  const sessionsMap = {};
  const lignesEvaluations = notesAffichees
    .map((n) => {
      const f = n.filiere || "Sans filière";
      const m = n.module || "Sans module";
      const t = n.typeEvaluation || "Contrôle continu";
      const d = n.date || "";
      const key = d + "|||" + f + "|||" + m + "|||" + t;
      if (!sessionsMap[key]) sessionsMap[key] = [];
      sessionsMap[key].push(n);
      return { ...n, _filiere: f, _filiereCode: codeFiliere(f), _module: m, _type: t };
    })
    .sort(
      (a, b) =>
        a._filiereCode.localeCompare(b._filiereCode, "fr", { numeric: true, sensitivity: "base" }) ||
        a._module.localeCompare(b._module, "fr", { sensitivity: "base" }) ||
        a._type.localeCompare(b._type, "fr", { sensitivity: "base" }) ||
        (a.date || "").localeCompare(b.date || "")
    );
  const bandeFiliereEvalParId = {};
  const premiereLigneFiliereEvalId = {};
  const derniereLigneFiliereEvalId = {};
  const premiereLigneModuleEvalId = {};
  {
    let bande = 0,
      filierePrecedente = null,
      modulePrecedent = null;
    lignesEvaluations.forEach((e, idx) => {
      if (e._filiereCode !== filierePrecedente) {
        bande = 1 - bande;
        filierePrecedente = e._filiereCode;
        premiereLigneFiliereEvalId[e.id] = true;
        if (idx > 0) derniereLigneFiliereEvalId[lignesEvaluations[idx - 1].id] = true;
      }
      if (e._module !== modulePrecedent) {
        modulePrecedent = e._module;
        premiereLigneModuleEvalId[e.id] = true;
      }
      bandeFiliereEvalParId[e.id] = bande;
    });
    if (lignesEvaluations.length > 0)
      derniereLigneFiliereEvalId[lignesEvaluations[lignesEvaluations.length - 1].id] = true;
  }
  return (
    /* @__PURE__ */ <div>
      <PageHeader
        title={role === "etudiant" ? "Mes notes" : "Évaluations"}
        sub={
          role === "etudiant"
            ? `${mesNotes.length} note(s) enregistrée(s)`
            : `Liste SharePoint "${SP.lists.notes}" — ${data.notes.length} enregistrements`
        }
        action={canEdit && /* @__PURE__ */ <Btn onClick={openAdd}>+ Ajouter une évaluation</Btn>}
      />
      {role === "etudiant" && monEtudiant && (
        /* @__PURE__ */ <div style={{ marginBottom: 20 }}>
          {[...new Set(mesNotes.map((n) => n.semestre).filter(Boolean))]
            .sort((a, b) => a.localeCompare(b, "fr", { numeric: true, sensitivity: "base" }))
            .map((sem) => {
              const mesPresencesEtudiant = (data.presences || []).filter(
                (p) =>
                  (user?.numInscription && p.numInscription === user.numInscription) ||
                  (user?.etudiantId && p.etudiantId === user.etudiantId)
              );
              const notesSem = mesNotes.filter((n) => n.semestre === sem);
              const vue = vueOuverte[sem];
              const basculer = (cible) => setVueOuverte((v) => ({ ...v, [sem]: v[sem] === cible ? null : cible }));
              return (
                /* @__PURE__ */ <div key={sem} style={{ marginBottom: 16 }}>
                  <div
                    style={{
                      display: "flex",
                      gap: 10,
                      flexWrap: "wrap",
                      alignItems: "center",
                      marginBottom: vue ? 12 : 0,
                    }}
                  >
                    <span style={{ fontSize: 12.5, fontWeight: 700, color: C.navy, minWidth: 90 }}>{sem} :</span>
                    <Btn onClick={() => basculer("synthetique")}>
                      📋 {vue === "synthetique" ? "Masquer" : "Voir"} le relevé synthétique
                    </Btn>
                    <Btn variant="light" onClick={() => basculer("detaille")}>
                      📋 {vue === "detaille" ? "Masquer" : "Voir"} le relevé détaillé
                    </Btn>
                  </div>
                  {vue === "synthetique" &&
                    (() => {
                      const { lignes, moyenneSemestre, tauxAbsencePriseEnCompte } = calculerLignesBulletin(
                        monEtudiant,
                        notesSem,
                        sem,
                        data.cours,
                        mesPresencesEtudiant,
                        data.configuration?.filieresAvecTauxAbsence
                      );
                      const fmt = (v) =>
                        v === null || v === void 0 || isNaN(v) ? "—" : v.toFixed(2).replace(".", ",");
                      const entetes = tauxAbsencePriseEnCompte
                        ? [
                            "Module",
                            "Moyenne CC/TP",
                            "Note EFM",
                            "Note Module",
                            "Taux d'absence atteint",
                            "Note rattrapage",
                            "Décision",
                          ]
                        : ["Module", "Moyenne CC/TP", "Note EFM", "Note Module", "Note rattrapage", "Décision"];
                      return (
                        /* @__PURE__ */ <Card style={{ padding: 20 }}>
                          <h4 style={{ margin: "0 0 4px", color: C.navy, fontSize: 15 }}>Relevé synthétique — {sem}</h4>
                          <div style={{ fontSize: 11.5, color: C.muted, marginBottom: 14 }}>
                            {monEtudiant.filiere} · {monEtudiant.niveau}
                          </div>
                          <div style={{ overflowX: "auto" }}>
                            <table style={{ width: "100%", borderCollapse: "collapse", minWidth: 640 }}>
                              <thead>
                                <tr style={{ background: "#000" }}>
                                  {entetes.map((h) => (
                                    /* @__PURE__ */ <th
                                      key={h}
                                      style={{
                                        padding: "9px 10px",
                                        color: "#fff",
                                        fontSize: 10.5,
                                        fontWeight: 700,
                                        textAlign: "left",
                                        whiteSpace: "nowrap",
                                      }}
                                    >
                                      {h}
                                    </th>
                                  ))}
                                </tr>
                              </thead>
                              <tbody>
                                {lignes.map((l, i) => (
                                  /* @__PURE__ */ <tr
                                    key={l.module}
                                    style={{
                                      background: i % 2 === 1 ? "#f5f5f5" : "#fff",
                                      borderBottom: `1px solid ${C.border}`,
                                    }}
                                  >
                                    <td
                                      style={{ padding: "8px 10px", fontSize: 12, color: "#8b4513", fontWeight: 600 }}
                                    >
                                      {l.module}
                                    </td>
                                    <td style={{ padding: "8px 10px", fontSize: 12 }}>{fmt(l.moyenneCC)}</td>
                                    <td style={{ padding: "8px 10px", fontSize: 12 }}>{fmt(l.efmNote)}</td>
                                    <td style={{ padding: "8px 10px", fontSize: 12, fontWeight: 700 }}>
                                      {fmt(l.noteModule)}
                                    </td>
                                    {tauxAbsencePriseEnCompte && (
                                      /* @__PURE__ */ <td style={{ padding: "8px 10px", fontSize: 12 }}>
                                        {l.tauxAbsence !== null ? `${l.tauxAbsence.toFixed(2)}%` : "—"}
                                      </td>
                                    )}
                                    <td style={{ padding: "8px 10px", fontSize: 12 }}>{fmt(l.rattrapageNote)}</td>
                                    <td style={{ padding: "8px 10px", fontSize: 12 }}>
                                      {l.decision && (
                                        /* @__PURE__ */ <Badge color={l.decision === "Validé" ? C.green : C.red}>
                                          {l.decision}
                                        </Badge>
                                      )}
                                    </td>
                                  </tr>
                                ))}
                              </tbody>
                            </table>
                          </div>
                          <div style={{ marginTop: 16, fontSize: 14, fontWeight: 700, color: C.navy }}>
                            Moyenne du semestre : <span style={{ fontSize: 16 }}>
                              {moyenneSemestre !== null ? fmt(moyenneSemestre) : "—"}
                            </span>
                          </div>
                        </Card>
                      );
                    })()}
                  {vue === "detaille" &&
                    (() => {
                      const fmt = (v) =>
                        v === null || v === void 0 || isNaN(v) ? "—" : v.toFixed(2).replace(".", ",");
                      const trie = [...notesSem].sort(
                        (a, b) =>
                          (b.date || "").localeCompare(a.date || "") ||
                          (a.module || "").localeCompare(b.module || "", "fr", { sensitivity: "base" })
                      );
                      return (
                        /* @__PURE__ */ <Card style={{ padding: 20 }}>
                          <h4 style={{ margin: "0 0 14px", color: C.navy, fontSize: 15 }}>
                            Relevé détaillé des évaluations — {sem}
                          </h4>
                          <div style={{ overflowX: "auto" }}>
                            <table style={{ width: "100%", borderCollapse: "collapse", minWidth: 500 }}>
                              <thead>
                                <tr style={{ background: `linear-gradient(135deg, ${C.navy} 0%, ${C.teal} 100%)` }}>
                                  {["Date", "Module", "Type d'évaluation", "Note /20"].map((h) => (
                                    /* @__PURE__ */ <th
                                      key={h}
                                      style={{
                                        padding: "9px 10px",
                                        color: "#fff",
                                        fontSize: 11,
                                        fontWeight: 700,
                                        textAlign: "left",
                                      }}
                                    >
                                      {h}
                                    </th>
                                  ))}
                                </tr>
                              </thead>
                              <tbody>
                                {trie.length === 0 ? (
                                  /* @__PURE__ */ <tr>
                                    <td
                                      colSpan={4}
                                      style={{ padding: 16, textAlign: "center", color: C.muted, fontSize: 12.5 }}
                                    >
                                      Aucune évaluation enregistrée
                                    </td>
                                  </tr>
                                ) : (
                                  trie.map((n, i) => {
                                    const [, col] = mention(n.note);
                                    return (
                                      /* @__PURE__ */ <tr
                                        key={n.id}
                                        style={{
                                          background: i % 2 === 1 ? C.greenLight : "#fff",
                                          borderBottom: `1px solid ${C.border}`,
                                        }}
                                      >
                                        <td style={{ padding: "8px 10px", fontSize: 12 }}>{formatDateFR(n.date)}</td>
                                        <td
                                          style={{
                                            padding: "8px 10px",
                                            fontSize: 12,
                                            color: "#8b4513",
                                            fontWeight: 600,
                                          }}
                                        >
                                          {n.module}
                                        </td>
                                        <td style={{ padding: "8px 10px", fontSize: 12 }}>
                                          <Badge
                                            color={
                                              n.typeEvaluation === "Rattrapage"
                                                ? C.purple
                                                : n.typeEvaluation === "Examen de fin de module"
                                                  ? C.orange
                                                  : C.teal
                                            }
                                          >
                                            {n.typeEvaluation}
                                          </Badge>
                                        </td>
                                        <td style={{ padding: "8px 10px", fontSize: 13, fontWeight: 700, color: col }}>
                                          {fmt(n.note)}
                                        </td>
                                      </tr>
                                    );
                                  })
                                )}
                              </tbody>
                            </table>
                          </div>
                        </Card>
                      );
                    })()}
                </div>
              );
            })}
        </div>
      )}
      {role !== "etudiant" && (
        /* @__PURE__ */ <Card style={{ padding: 16, marginBottom: 16 }}>
          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(160px, 1fr))", gap: 10 }}>
            <input
              placeholder="🔍  Nom ou N° inscription..."
              value={searchListe}
              onChange={(e) => setSearchListe(e.target.value)}
              style={{
                padding: "9px 12px",
                border: `1px solid ${C.border}`,
                borderRadius: 8,
                fontSize: 13,
                boxSizing: "border-box",
                fontFamily: "inherit",
              }}
            />
            <select
              value={filtreFiliereListe}
              onChange={(e) => setFiltreFiliereListe(e.target.value)}
              style={{
                padding: "9px 12px",
                border: `1px solid ${C.border}`,
                borderRadius: 8,
                fontSize: 13,
                fontFamily: "inherit",
                background: C.white,
              }}
            >
              <option value="">-- Tous les groupes --</option>
              {filieresOptions.map((f) => (
                /* @__PURE__ */ <option key={f} value={f}>
                  {f}
                </option>
              ))}
            </select>
            <select
              value={filtreSemestreListe}
              onChange={(e) => setFiltreSemestreListe(e.target.value)}
              style={{
                padding: "9px 12px",
                border: `1px solid ${C.border}`,
                borderRadius: 8,
                fontSize: 13,
                fontFamily: "inherit",
                background: C.white,
              }}
            >
              <option value="">-- Tous les semestres --</option>
              {semestresOptions.map((s) => (
                /* @__PURE__ */ <option key={s} value={s}>
                  {s}
                </option>
              ))}
            </select>
            <select
              value={filtreModuleListe}
              onChange={(e) => setFiltreModuleListe(e.target.value)}
              style={{
                padding: "9px 12px",
                border: `1px solid ${C.border}`,
                borderRadius: 8,
                fontSize: 13,
                fontFamily: "inherit",
                background: C.white,
              }}
            >
              <option value="">-- Tous les modules --</option>
              {modulesOptionsListe.map((m) => (
                /* @__PURE__ */ <option key={m} value={m}>
                  {m}
                </option>
              ))}
            </select>
            <select
              value={filtreTypeListe}
              onChange={(e) => setFiltreTypeListe(e.target.value)}
              style={{
                padding: "9px 12px",
                border: `1px solid ${C.border}`,
                borderRadius: 8,
                fontSize: 13,
                fontFamily: "inherit",
                background: C.white,
              }}
            >
              <option value="">-- Tous les types --</option>
              {TYPES_EVALUATION.map((t) => (
                /* @__PURE__ */ <option key={t} value={t}>
                  {t}
                </option>
              ))}
            </select>
          </div>
          {canPrint && (
            /* @__PURE__ */ <div
              style={{
                display: "flex",
                gap: 10,
                flexWrap: "wrap",
                justifyContent: "center",
                marginTop: 12,
                paddingTop: 12,
                borderTop: `1px solid ${C.light}`,
              }}
            >
              <Btn variant="navy" onClick={() => imprimerRelevesGroupe("synthetique")}>
                🖨️ Imprimer les relevés synthétiques (résultats filtrés)
              </Btn>
              <Btn variant="navy" onClick={() => imprimerRelevesGroupe("detaille")}>
                🖨️ Imprimer les relevés détaillés (résultats filtrés)
              </Btn>
            </div>
          )}
        </Card>
      )}
      {lignesEvaluations.length === 0 ? (
        /* @__PURE__ */ <Card style={{ padding: 32, textAlign: "center", color: C.muted, fontSize: 13 }}>
          Aucune évaluation trouvée
        </Card>
      ) : (
        /* @__PURE__ */ <Card style={{ overflow: "hidden" }}>
          <DataTable
            cols={[
              { key: "_filiereCode", label: "Filière", render: (v) => /* @__PURE__ */ <FiliereBadge code={v} /> },
              { key: "date", label: "Date", render: (v) => formatDateFR(v) || v },
              {
                key: "_module",
                label: "Module",
                render: (v, row) => (
                  /* @__PURE__ */ <span
                    style={{
                      fontWeight: 700,
                      color: "#8b4513",
                      display: "inline-block",
                      borderTop: premiereLigneModuleEvalId[row.id] ? `2px dashed ${C.orange}` : void 0,
                      paddingTop: premiereLigneModuleEvalId[row.id] ? 4 : 0,
                    }}
                  >
                    {v}
                  </span>
                ),
              },
              {
                key: "_type",
                label: "Type d'évaluation",
                render: (v, row) => (
                  /* @__PURE__ */ <div
                    style={{
                      borderTop: premiereLigneModuleEvalId[row.id] ? `2px dashed ${C.orange}` : void 0,
                      paddingTop: premiereLigneModuleEvalId[row.id] ? 4 : 0,
                      display: "inline-block",
                    }}
                  >
                    <Badge color={v === "Rattrapage" ? C.purple : v === "Examen de fin de module" ? C.orange : C.teal}>
                      {v}
                    </Badge>
                  </div>
                ),
              },
              {
                key: "etudiant",
                label: "Étudiant",
                render: (v, row) => (
                  /* @__PURE__ */ <div
                    style={{
                      borderTop: premiereLigneModuleEvalId[row.id] ? `2px dashed ${C.orange}` : void 0,
                      paddingTop: premiereLigneModuleEvalId[row.id] ? 4 : 0,
                    }}
                  >
                    {v}
                  </div>
                ),
              },
              {
                key: "numInscription",
                label: "N° Inscription",
                render: (v, row) => (
                  /* @__PURE__ */ <div
                    style={{
                      borderTop: premiereLigneModuleEvalId[row.id] ? `2px dashed ${C.orange}` : void 0,
                      paddingTop: premiereLigneModuleEvalId[row.id] ? 4 : 0,
                    }}
                  >
                    {v}
                  </div>
                ),
              },
              {
                key: "note",
                label: "Note /20",
                render: (v, row) => {
                  const [, col] = mention(v);
                  return (
                    /* @__PURE__ */ <div
                      style={{
                        borderTop: premiereLigneModuleEvalId[row.id] ? `2px dashed ${C.orange}` : void 0,
                        paddingTop: premiereLigneModuleEvalId[row.id] ? 4 : 0,
                      }}
                    >
                      <span style={{ fontWeight: 700, color: col, fontSize: 14 }}>{v}</span>
                    </div>
                  );
                },
              },
              ...(canEdit
                ? [
                    {
                      key: "_modifier",
                      label: "Modifier",
                      render: (_, row) => (
                        /* @__PURE__ */ <Btn
                          small={true}
                          variant="navy"
                          onClick={() =>
                            openEditSession(
                              row.date,
                              row._filiere,
                              row.semestre,
                              row._module,
                              row._type,
                              sessionsMap[row.date + "|||" + row._filiere + "|||" + row._module + "|||" + row._type]
                            )
                          }
                        >
                          ✏️
                        </Btn>
                      ),
                    },
                  ]
                : []),
              ...(canEdit
                ? [
                    {
                      key: "_suppr",
                      label: "Suppr.",
                      render: (_, row) => (
                        /* @__PURE__ */ <Btn small={true} variant="danger" onClick={() => handleDeleteNote(row)}>
                          🗑
                        </Btn>
                      ),
                    },
                  ]
                : []),
            ]}
            rows={lignesEvaluations}
            emptyMsg="Aucune évaluation trouvée"
            compact={true}
            rowStyle={(row) => ({
              background: bandeFiliereEvalParId[row.id] === 0 ? C.navyLight : "#e0f4f2",
              borderTop: premiereLigneFiliereEvalId[row.id] ? `2.5px solid ${C.navy}` : void 0,
              borderBottom: derniereLigneFiliereEvalId[row.id] ? `2.5px solid ${C.navy}` : `1px solid ${C.border}`,
            })}
          />
        </Card>
      )}
      {showModal && (
        /* @__PURE__ */ <Modal
          title={
            Object.keys(notesIdsEdition).length > 0
              ? "Modifier l'évaluation → SharePoint"
              : "Ajouter une évaluation → SharePoint"
          }
          onClose={() => setShowModal(false)}
          width={640}
        >
          <div className="eco-grid2" style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "0 14px" }}>
            <Input
              label="Date de l'évaluation *"
              type="date"
              value={form.date}
              min={semestreInfoForm?.dateDebut || void 0}
              max={semestreInfoForm?.dateFin || void 0}
              onChange={(e) => setForm((f) => ({ ...f, date: e.target.value }))}
            />
            <CreatableSelect
              label="Semestre *"
              value={form.semestre}
              onChange={(v) => {
                const sInfo = (data.semestresDates || []).find((s) => s.semestre === v);
                setForm((f) => ({ ...f, semestre: v, module: "", date: sInfo?.dateDebut || f.date }));
              }}
              options={semestresOptions}
            />
          </div>
          {semestreInfoForm && (semestreInfoForm.dateDebut || semestreInfoForm.dateFin) && (
            /* @__PURE__ */ <div style={{ fontSize: 11, color: C.muted, marginTop: -10, marginBottom: 13 }}>
              📅 Période du semestre (Paramétrage) : {formatDateFR(semestreInfoForm.dateDebut) || "—"} → {formatDateFR(semestreInfoForm.dateFin) || "—"}
            </div>
          )}
          <Sel
            label="Filière *"
            value={form.filiere}
            onChange={(e) => {
              setForm((f) => ({ ...f, filiere: e.target.value, module: "" }));
              setNotesSaisies({});
            }}
          >
            <option value="">-- Sélectionner --</option>
            {filieresOptions.map((f) => (
              /* @__PURE__ */ <option key={f} value={f}>
                {f}
              </option>
            ))}
          </Sel>
          <Sel
            label="Module *"
            value={form.module}
            onChange={(e) => setForm((f) => ({ ...f, module: e.target.value }))}
          >
            <option value="">-- Sélectionner --</option>
            {modulesOptions.map((c) => (
              /* @__PURE__ */ <option key={c.id} value={c.intitule}>
                {c.intitule}
              </option>
            ))}
          </Sel>
          <Sel
            label="Type d'évaluation *"
            value={form.typeEvaluation}
            onChange={(e) => setForm((f) => ({ ...f, typeEvaluation: e.target.value }))}
          >
            {TYPES_EVALUATION.map((t) => (
              /* @__PURE__ */ <option key={t}>{t}</option>
            ))}
          </Sel>
          {form.filiere && (
            /* @__PURE__ */ <div style={{ marginTop: 6, marginBottom: 16 }}>
              <div
                style={{
                  fontSize: 11.5,
                  fontWeight: 700,
                  color: C.navy,
                  textTransform: "uppercase",
                  letterSpacing: "0.06em",
                  marginBottom: 8,
                }}
              >
                Saisir les notes ({etudiantsFiliere.length} étudiant(s) dans ce groupe
                {nbNotesSaisies > 0 ? ` — ${nbNotesSaisies} saisie(s)` : ""})
              </div>
              {etudiantsFiliere.length === 0 ? (
                /* @__PURE__ */ <div
                  style={{
                    textAlign: "center",
                    padding: 16,
                    color: C.muted,
                    fontSize: 12.5,
                    border: `1px solid ${C.border}`,
                    borderRadius: 8,
                  }}
                >
                  Aucun étudiant dans ce groupe
                </div>
              ) : (
                /* @__PURE__ */ <div
                  style={{ border: `1px solid ${C.border}`, borderRadius: 10, maxHeight: 280, overflowY: "auto" }}
                >
                  {etudiantsFiliere.map((et, i) => {
                    const valeur = notesSaisies[et.id] || "";
                    const invalide =
                      valeur !== "" && (isNaN(parseFloat(valeur)) || parseFloat(valeur) < 0 || parseFloat(valeur) > 20);
                    return (
                      /* @__PURE__ */ <div
                        key={et.id}
                        style={{
                          display: "flex",
                          alignItems: "center",
                          gap: 10,
                          padding: "8px 14px",
                          borderBottom: i < etudiantsFiliere.length - 1 ? `1px solid ${C.light}` : "none",
                          background: i % 2 === 0 ? C.white : C.light,
                        }}
                      >
                        <span style={{ fontSize: 13, fontWeight: 600, flex: 1, minWidth: 0 }}>
                          {et.prenom} {et.nom}
                        </span>
                        <span style={{ fontSize: 11, color: C.muted }}>{et.numInscription}</span>
                        <input
                          type="number"
                          min="0"
                          max="20"
                          step="0.25"
                          placeholder="Note /20"
                          value={valeur}
                          onChange={(e) => setNotesSaisies((n) => ({ ...n, [et.id]: e.target.value }))}
                          style={{
                            width: 90,
                            padding: "6px 8px",
                            border: `1px solid ${invalide ? C.red : C.border}`,
                            borderRadius: 6,
                            fontSize: 13,
                            fontFamily: "inherit",
                            textAlign: "center",
                            background: invalide ? "#fdeeec" : C.white,
                          }}
                        />
                      </div>
                    );
                  })}
                </div>
              )}
              {Object.values(notesSaisies).some(
                (v) => v !== "" && (isNaN(parseFloat(v)) || parseFloat(v) < 0 || parseFloat(v) > 20)
              ) && (
                /* @__PURE__ */ <div style={{ fontSize: 11.5, color: C.red, fontWeight: 700, marginTop: 6 }}>
                  ⚠️ Certaines notes sont hors de la plage 0-20 — corrigez-les avant d'enregistrer.
                </div>
              )}
            </div>
          )}
          <div style={{ display: "flex", gap: 10, justifyContent: "flex-end" }}>
            <Btn
              variant="light"
              onClick={() => {
                setShowModal(false);
                setNotesIdsEdition({});
              }}
            >
              Annuler
            </Btn>
            <Btn
              onClick={handleSaveNotes}
              disabled={!form.date || !form.semestre || !form.filiere || !form.module || nbNotesSaisies === 0 || saving}
            >
              {saving
                ? "Enregistrement..."
                : Object.keys(notesIdsEdition).length > 0
                  ? `Mettre à jour ${nbNotesSaisies > 0 ? `(${nbNotesSaisies})` : ""} → SharePoint`
                  : `Enregistrer ${nbNotesSaisies > 0 ? `(${nbNotesSaisies})` : ""} → SharePoint`}
            </Btn>
          </div>
        </Modal>
      )}
    </div>
  );
}
const SEANCES_HORAIRE = ["09h00 - 10h30", "10h45 - 12h15", "14h00 - 15h30", "15h45 - 17h15", "19h00 - 21h00"];
const SEANCE_DUREE = {
  "09h00 - 10h30": 1.5,
  "10h45 - 12h15": 1.5,
  "14h00 - 15h30": 1.5,
  "15h45 - 17h15": 1.5,
  "19h00 - 21h00": 2,
};
function PieChart({ data, size = 190 }) {
  const total = data.reduce((s, d) => s + d.value, 0);
  const radius = size / 2;
  const cx = radius,
    cy = radius;
  if (total <= 0) {
    return (
      /* @__PURE__ */ <svg width={size} height={size} viewBox={`0 0 ${size} ${size}`}>
        <circle cx={cx} cy={cy} r={radius - 2} fill={C.light} stroke={C.border} />
      </svg>
    );
  }
  const toRad = (a) => (a * Math.PI) / 180;
  const nonZero = data.filter((d) => d.value > 0);
  let cumulative = -90;
  const infos = nonZero.map((d) => {
    const pct = (d.value / total) * 100;
    const angle = (d.value / total) * 360;
    const startAngle = cumulative;
    const endAngle = cumulative + angle;
    const midAngle = (startAngle + endAngle) / 2;
    cumulative = endAngle;
    return { ...d, pct, startAngle, endAngle, midAngle };
  });
  const slices = infos.map((d, i) => {
    if (nonZero.length === 1) {
      return (
        /* @__PURE__ */ <circle
          key={`s${i}`}
          cx={cx}
          cy={cy}
          r={radius - 2}
          fill={d.color}
          stroke="white"
          strokeWidth="2"
        />
      );
    }
    const x1 = cx + (radius - 2) * Math.cos(toRad(d.startAngle));
    const y1 = cy + (radius - 2) * Math.sin(toRad(d.startAngle));
    const x2 = cx + (radius - 2) * Math.cos(toRad(d.endAngle));
    const y2 = cy + (radius - 2) * Math.sin(toRad(d.endAngle));
    const largeArc = d.endAngle - d.startAngle > 180 ? 1 : 0;
    const path = `M ${cx} ${cy} L ${x1} ${y1} A ${radius - 2} ${radius - 2} 0 ${largeArc} 1 ${x2} ${y2} Z`;
    return /* @__PURE__ */ <path key={`s${i}`} d={path} fill={d.color} stroke="white" strokeWidth="2" />;
  });
  const labels = infos
    .filter((d) => d.pct >= 4 && !d.noLabel)
    .map((d, i) => {
      const labelRadius = (radius - 2) * 0.66;
      const lx = cx + labelRadius * Math.cos(toRad(d.midAngle));
      const ly = cy + labelRadius * Math.sin(toRad(d.midAngle));
      const valeurAffichee = d.taux !== void 0 && d.taux !== null ? d.taux.toFixed(2) : Math.round(d.pct);
      return (
        /* @__PURE__ */ <text
          key={`l${i}`}
          x={lx}
          y={ly}
          textAnchor="middle"
          dominantBaseline="middle"
          style={{
            fontSize: 12.5,
            fontWeight: 800,
            fill: "#ffffff",
            paintOrder: "stroke",
            stroke: "rgba(0,0,0,0.45)",
            strokeWidth: 3,
          }}
        >
          {valeurAffichee}%
        </text>
      );
    });
  return (
    /* @__PURE__ */ <svg width={size} height={size} viewBox={`0 0 ${size} ${size}`}>
      {slices}
      {labels}
    </svg>
  );
}
function Presences({ data, setData, role, user, toast }) {
  const canEdit = ["administrateur", "professeur"].includes(role);
  const [showModal, setShowModal] = useState(false);
  const [saving, setSaving] = useState(false);
  const emptyForm = {
    date: today(),
    professeur: "",
    professeurCin: "",
    module: "",
    seance: "",
    filiere: "",
    semestre: "",
  };
  const [form, setForm] = useState(emptyForm);
  const [absents, setAbsents] = useState({});
  const [expandedNodes, setExpandedNodes] = useState({});
  const toggleNode = (key) => setExpandedNodes((e) => ({ ...e, [key]: !e[key] }));
  const [showJustif, setShowJustif] = useState(null);
  const [justifForm, setJustifForm] = useState({});
  const [filtreRecherche, setFiltreRecherche] = useState("");
  const [filtreFiliereListe, setFiltreFiliereListe] = useState("");
  const [filtreSemestreListe, setFiltreSemestreListe] = useState("");
  const [filtreModuleListe, setFiltreModuleListe] = useState("");
  const [filtreProfesseurListe, setFiltreProfesseurListe] = useState("");
  const [filtreStatutListe, setFiltreStatutListe] = useState("");
  const [filtreDateDebut, setFiltreDateDebut] = useState("");
  const [filtreDateFin, setFiltreDateFin] = useState("");
  const monProfesseur = (data.professeurs || []).find(
    (p) => p.email && user?.email && p.email.toLowerCase() === user.email.toLowerCase()
  );
  const mesPresences =
    role === "etudiant"
      ? (data.presences || []).filter(
          (p) =>
            (user?.numInscription && p.numInscription === user.numInscription) ||
            (user?.etudiantId && p.etudiantId === user.etudiantId)
        )
      : role === "professeur"
        ? (data.presences || []).filter((p) => p.enseignantCin === monProfesseur?.cin)
        : data.presences || [];
  const monEtudiant = (data.etudiants || []).find(
    (e) =>
      (user?.numInscription && e.numInscription === user.numInscription) ||
      (user?.etudiantId && e.etudiantId === user.etudiantId)
  );
  const openJustif = (p) => {
    setShowJustif(p);
    setJustifForm({
      statut: p.statut || "Absent",
      documentJustification: p.documentJustification || "",
      justifiePar: p.justifiePar || "",
      dateJustification: p.dateJustification || today(),
    });
  };
  const handleSaveJustif = () => {
    if (!showJustif) return;
    const updated = { ...showJustif, ...justifForm };
    setData((d) => ({ ...d, presences: d.presences.map((x) => (x.id === showJustif.id ? updated : x)) }));
    if (user?.token && showJustif._spId) {
      saveToSP(
        "ECOGEST_Presences",
        {
          Statut: justifForm.statut || "Absent",
          DocumentJustification: justifForm.documentJustification || "",
          JustifiePar: justifForm.justifiePar || "",
          DateJustification: justifForm.dateJustification || "",
        },
        user.token,
        showJustif._spId
      ).then((result) => {
        if (!result) {
          toast("⚠️ Échec de la mise à jour dans SharePoint");
          return;
        }
        toast("Justification enregistrée dans SharePoint ✓");
      });
    } else {
      toast("⚠️ Non connecté à SharePoint — justification non enregistrée");
    }
    setShowJustif(null);
  };
  const presents = mesPresences.filter((p) => p.statut === "Présent").length;
  const absentsCount = mesPresences.filter((p) => p.statut === "Absent").length;
  const justifie = mesPresences.filter((p) => p.statut === "Justifié").length;
  const professeursRecordsTous = [...(data.professeurs || [])].sort((a, b) =>
    `${a.prenom || ""} ${a.nom || ""}`
      .trim()
      .localeCompare(`${b.prenom || ""} ${b.nom || ""}`.trim(), "fr", { sensitivity: "base" })
  );
  const professeursRecords = role === "professeur" && monProfesseur ? [monProfesseur] : professeursRecordsTous;
  const semestresOptions = [
    .../* @__PURE__ */ new Set([
      ...SEMESTRES_MODULE_SEED,
      ...(data.cours || []).map((c) => c.semestre).filter(Boolean),
    ]),
  ].sort((a, b) => a.localeCompare(b, "fr", { sensitivity: "base" }));
  const modulesOptions = (data.cours || []).filter(
    (c) =>
      (role === "professeur"
        ? c.professeurCin === monProfesseur?.cin
        : !form.professeurCin || c.professeurCin === form.professeurCin) &&
      (!form.filiere || (c.filieresListe || [c.filiere]).includes(form.filiere)) &&
      (!form.semestre || c.semestre === form.semestre)
  );
  const semestreInfoForm = (data.semestresDates || []).find((s) => s.semestre === form.semestre);
  const filieresOptions = getFilieres(data);
  const etudiantsFiliere = form.filiere ? data.etudiants.filter((e) => e.filiere === form.filiere) : [];
  const lundiDeLaSemaine = (dateStr) => {
    const [an, mois, jourNum] = dateStr.split("-").map(Number);
    const d = new Date(an, mois - 1, jourNum);
    const jour = d.getDay();
    const diff = jour === 0 ? -6 : 1 - jour;
    d.setDate(d.getDate() + diff);
    return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
  };
  const [ficheFiliere, setFicheFiliere] = useState("");
  const [ficheSemaine, setFicheSemaine] = useState(lundiDeLaSemaine(today()));
  const handleImprimerFicheHebdo = () => {
    if (!ficheFiliere) {
      toast("⚠️ Sélectionnez un groupe pour la fiche de présence");
      return;
    }
    const filiereObj = (data.filieres || []).find((f) => f.intitule === ficheFiliere);
    const etudiantsGroupe = (data.etudiants || []).filter((e) => e.filiere === ficheFiliere && e.statut !== "Inactif");
    genererFichePresenceHebdo(filiereObj, etudiantsGroupe, lundiDeLaSemaine(ficheSemaine), [
      "Lundi",
      "Mardi",
      "Mercredi",
      "Jeudi",
      "Vendredi",
    ]);
  };
  const openAdd = () => {
    setForm({
      ...emptyForm,
      professeur: role === "professeur" && monProfesseur ? `${monProfesseur.prenom} ${monProfesseur.nom}` : "",
      professeurCin: role === "professeur" && monProfesseur ? monProfesseur.cin || "" : "",
    });
    setAbsents({});
    setShowModal(true);
  };
  const toggleAbsent = (etId) => setAbsents((a) => ({ ...a, [etId]: !a[etId] }));
  const nbCoches = Object.values(absents).filter(Boolean).length;
  const handleSave = async () => {
    const etudiantsAbsents = etudiantsFiliere.filter((e) => absents[e.id]);
    if (etudiantsAbsents.length === 0) {
      toast("⚠️ Cochez au moins un étudiant absent");
      return;
    }
    if (!user?.token) {
      toast("⚠️ Non connecté à SharePoint — absences non enregistrées");
      return;
    }
    setSaving(true);
    const results = await Promise.all(
      etudiantsAbsents.map((et) => {
        const spF = {
          Title: `${et.nom} ${et.prenom} - ${form.date}`,
          EtudiantId: String(et.id),
          NomEtudiant: `${et.nom} ${et.prenom}`,
          NumInscription: et.numInscription || "",
          Module: form.module || "",
          DatePresence: form.date || "",
          Heure: form.seance || "",
          Statut: "Absent",
          Enseignant: form.professeur || "",
          EnseignantCIN: form.professeurCin || "",
          Filiere: form.filiere || "",
          Seance: form.seance || "",
          Semestre: form.semestre || "",
        };
        return saveToSP("ECOGEST_Presences", spF, user.token).then((result) => (result ? { et, result } : null));
      })
    );
    setSaving(false);
    const ok = results.filter(Boolean);
    if (ok.length > 0) {
      setData((d) => ({
        ...d,
        presences: [
          ...d.presences,
          ...ok.map(({ et, result }) => ({
            id: Date.now() + Math.random(),
            _spId: result.id,
            etudiant: `${et.prenom} ${et.nom}`,
            etudiantId: et.id,
            numInscription: et.numInscription,
            module: form.module,
            date: form.date,
            heure: form.seance,
            statut: "Absent",
            enseignant: form.professeur,
            enseignantCin: form.professeurCin,
            filiere: form.filiere,
            seance: form.seance,
            semestre: form.semestre,
          })),
        ],
      }));
    }
    if (ok.length < etudiantsAbsents.length)
      toast(`⚠️ ${etudiantsAbsents.length - ok.length} absence(s) non enregistrée(s) — échec SharePoint`);
    else toast(`${ok.length} absence(s) enregistrée(s) dans SharePoint ✓`);
    if (ok.length > 0) {
      setShowModal(false);
      setForm(emptyForm);
      setAbsents({});
    }
  };
  return (
    /* @__PURE__ */ <div>
      <PageHeader
        title={role === "etudiant" ? "Mes absences" : "Gestion des absences"}
        sub={
          role === "etudiant"
            ? `${mesPresences.length} absence(s) enregistrée(s)`
            : `Liste SharePoint "${SP.lists.presences}" — ${data.presences.length} enregistrements`
        }
        action={
          canEdit ? (
            /* @__PURE__ */ <Btn onClick={openAdd}>+ Enregistrer des absences</Btn>
          ) : role === "etudiant" && monEtudiant ? (
            /* @__PURE__ */ <Btn onClick={() => genererReleveAbsences(monEtudiant, mesPresences, data.cours)}>
              📄 Télécharger mon relevé d'absences (PDF)
            </Btn>
          ) : null
        }
      />
      {canEdit && (
        /* @__PURE__ */ <Card style={{ padding: 16, marginBottom: 16 }}>
          <div style={{ fontSize: 12.5, fontWeight: 700, color: C.navy, marginBottom: 10 }}>
            🖨️ Fiche de présence hebdomadaire (à imprimer et remplir manuellement)
          </div>
          <div style={{ display: "flex", gap: 10, flexWrap: "wrap", alignItems: "flex-end" }}>
            <div style={{ flex: 1, minWidth: 220 }}>
              <label style={{ fontSize: 11, fontWeight: 600, color: "#555", display: "block", marginBottom: 4 }}>
                Groupe
              </label>
              <select
                value={ficheFiliere}
                onChange={(e) => setFicheFiliere(e.target.value)}
                style={{
                  width: "100%",
                  padding: "9px 12px",
                  border: `1px solid ${C.border}`,
                  borderRadius: 8,
                  fontSize: 13,
                  fontFamily: "inherit",
                  background: C.white,
                }}
              >
                <option value="">-- Sélectionner --</option>
                {filieresOptions.map((f) => (
                  /* @__PURE__ */ <option key={f} value={f}>
                    {f}
                  </option>
                ))}
              </select>
            </div>
            <div>
              <label style={{ fontSize: 11, fontWeight: 600, color: "#555", display: "block", marginBottom: 4 }}>
                Semaine du (lundi)
              </label>
              <input
                type="date"
                value={ficheSemaine}
                onChange={(e) => setFicheSemaine(e.target.value)}
                style={{
                  padding: "9px 12px",
                  border: `1px solid ${C.border}`,
                  borderRadius: 8,
                  fontSize: 13,
                  fontFamily: "inherit",
                }}
              />
            </div>
            <Btn onClick={handleImprimerFicheHebdo}>🖨️ Imprimer la fiche de présences</Btn>
          </div>
        </Card>
      )}
      {role === "etudiant" &&
        (() => {
          const PALETTE = [C.navy, C.orange, C.teal, C.purple, C.green, C.red, "#c99a2e", "#6b7280"];
          const parSemestreModule = {};
          mesPresences.forEach((p) => {
            const mod = p.module || "Module non précisé";
            const moduleInfo = (data.cours || []).find((c) => c.intitule === mod);
            const sem = moduleInfo?.semestre || p.semestre || "Semestre non précisé";
            if (!parSemestreModule[sem]) parSemestreModule[sem] = {};
            if (!parSemestreModule[sem][mod]) parSemestreModule[sem][mod] = 0;
            parSemestreModule[sem][mod] += SEANCE_DUREE[p.seance || p.heure] || 1.5;
          });
          const semestreKeys = Object.keys(parSemestreModule).sort((a, b) =>
            a.localeCompare(b, "fr", { numeric: true, sensitivity: "base" })
          );
          if (semestreKeys.length === 0) return null;
          return (
            /* @__PURE__ */ <Card style={{ padding: 24, marginBottom: 20 }}>
              <h4 style={{ margin: "0 0 16px", color: C.navy, fontSize: 14.5 }}>
                📊 Détail du taux d'absence par module
              </h4>
              <div style={{ display: "flex", flexDirection: "row", flexWrap: "wrap", gap: 24 }}>
                {semestreKeys.map((sem) => {
                  const parModule = parSemestreModule[sem];
                  const moduleKeys = Object.keys(parModule).sort((a, b) =>
                    a.localeCompare(b, "fr", { sensitivity: "base" })
                  );
                  const analyse = moduleKeys.map((m, i) => {
                    const heuresManquees = parModule[m];
                    const moduleInfo = (data.cours || []).find((c) => c.intitule === m);
                    const heuresTotal = moduleInfo?.heures || 0;
                    const taux = heuresTotal > 0 ? Math.round((heuresManquees / heuresTotal) * 10000) / 100 : null;
                    return { module: m, heuresManquees, heuresTotal, taux, color: PALETTE[i % PALETTE.length] };
                  });
                  const pieData = analyse.map((a) => ({
                    label: a.module,
                    value: a.heuresManquees,
                    color: a.color,
                    taux: a.taux,
                  }));
                  return (
                    /* @__PURE__ */ <div key={sem} style={{ flex: "1 1 340px", minWidth: 300 }}>
                      <div
                        style={{
                          fontSize: 12.5,
                          fontWeight: 800,
                          color: C.orange,
                          textTransform: "uppercase",
                          letterSpacing: "0.05em",
                          marginBottom: 12,
                        }}
                      >
                        {sem}
                      </div>
                      <div style={{ display: "flex", flexDirection: "column", alignItems: "center", gap: 16 }}>
                        <PieChart data={pieData} size={160} />
                        <div style={{ width: "100%" }}>
                          {analyse.map((a) => (
                            /* @__PURE__ */ <div
                              key={a.module}
                              style={{
                                display: "flex",
                                alignItems: "center",
                                gap: 10,
                                padding: "9px 0",
                                borderBottom: `1px solid ${C.light}`,
                              }}
                            >
                              <span
                                style={{ width: 12, height: 12, borderRadius: 3, background: a.color, flexShrink: 0 }}
                              />
                              <div style={{ flex: 1, minWidth: 0 }}>
                                <div style={{ fontSize: 12.5, fontWeight: 600, color: "#232733" }}>{a.module}</div>
                                <div style={{ fontSize: 11, color: C.muted }}>
                                  {a.heuresManquees}h manquée(s)
                                  {a.heuresTotal > 0
                                    ? ` sur ${a.heuresTotal}h`
                                    : " — volume horaire du module non renseigné"}
                                </div>
                              </div>
                              {a.taux !== null && (
                                /* @__PURE__ */ <span
                                  style={{
                                    fontSize: 13,
                                    fontWeight: 800,
                                    color: a.taux >= 20 ? C.red : a.taux >= 10 ? C.orange : C.green,
                                    flexShrink: 0,
                                  }}
                                >
                                  {a.taux.toFixed(2)}%
                                </span>
                              )}
                            </div>
                          ))}
                        </div>
                      </div>
                    </div>
                  );
                })}
              </div>
            </Card>
          );
        })()}
      {role === "etudiant" &&
        (() => {
          const mesModules = (data.cours || []).filter(
            (c) => !monEtudiant || (c.filieresListe || [c.filiere]).includes(monEtudiant.filiere)
          );
          const parSemestre = {};
          mesModules.forEach((c) => {
            const key = c.semestre || "Sans semestre";
            if (!parSemestre[key]) parSemestre[key] = [];
            parSemestre[key].push(c);
          });
          const semestreKeys = Object.keys(parSemestre).sort((a, b) =>
            a.localeCompare(b, "fr", { numeric: true, sensitivity: "base" })
          );
          if (semestreKeys.length === 0) return null;
          const cols = [
            { key: "intitule", label: "Module" },
            { key: "professeur", label: "Professeur", render: (v) => (role === "etudiant" ? masquerCinProf(v) : v) },
            { key: "heures", label: "Vol. horaire", render: (v) => `${v || 0}h` },
            {
              key: "_absences",
              label: "Absences",
              render: (_, row) => {
                const nb = mesPresences.filter((p) => p.module === row.intitule).length;
                return (
                  /* @__PURE__ */ <span
                    style={{
                      display: "inline-flex",
                      alignItems: "center",
                      gap: 3,
                      background: nb > 0 ? "#fdeeec" : C.navyLight,
                      color: nb > 0 ? C.red : C.navy,
                      fontSize: 10,
                      fontWeight: 700,
                      padding: "2px 7px",
                      borderRadius: 16,
                    }}
                  >
                    ❌ {nb}
                  </span>
                );
              },
            },
            {
              key: "_taux",
              label: "Synthèse",
              render: (_, row) => {
                const heuresManquees = mesPresences
                  .filter((p) => p.module === row.intitule)
                  .reduce((s, p) => s + (SEANCE_DUREE[p.seance || p.heure] || 1.5), 0);
                const heuresTotal = row.heures || 0;
                const taux = heuresTotal > 0 ? Math.round((heuresManquees / heuresTotal) * 10000) / 100 : 0;
                const barColor = taux >= 20 ? C.red : taux >= 10 ? C.orange : C.green;
                return (
                  /* @__PURE__ */ <div style={{ minWidth: 160 }}>
                    <div style={{ display: "flex", gap: 6, alignItems: "center", flexWrap: "wrap", marginBottom: 5 }}>
                      <span
                        style={{
                          background: C.navyLight,
                          color: C.navy,
                          border: `1px solid ${C.navy}40`,
                          borderRadius: 16,
                          padding: "1px 7px",
                          fontSize: 9.5,
                          fontWeight: 600,
                        }}
                      >
                        {heuresTotal}h total
                      </span>
                      <span
                        style={{
                          background: barColor + "18",
                          color: barColor,
                          border: `1px solid ${barColor}40`,
                          borderRadius: 16,
                          padding: "1px 7px",
                          fontSize: 9.5,
                          fontWeight: 600,
                        }}
                      >
                        {heuresManquees}h manquée(s)
                      </span>
                    </div>
                    <div style={{ display: "flex", alignItems: "center", gap: 9 }}>
                      <div
                        style={{
                          flex: 1,
                          minWidth: 70,
                          background: "#e7ebf2",
                          borderRadius: 20,
                          height: 12,
                          boxShadow: "inset 0 1px 4px rgba(20,30,60,0.22)",
                          overflow: "hidden",
                        }}
                      >
                        <div
                          style={{
                            width: `${Math.min(taux, 100)}%`,
                            height: "100%",
                            background: `linear-gradient(90deg, ${barColor}bb 0%, ${barColor} 100%)`,
                            borderRadius: 20,
                            transition: "width 0.5s",
                            boxShadow: `0 0 7px ${barColor}99`,
                          }}
                        />
                      </div>
                      <span
                        style={{ fontSize: 12.5, fontWeight: 800, color: barColor, minWidth: 46, textAlign: "right" }}
                      >
                        {taux.toFixed(2)}%
                      </span>
                    </div>
                  </div>
                );
              },
            },
          ];
          return semestreKeys.map((sk) => (
            /* @__PURE__ */ <div key={sk} style={{ marginBottom: 14 }}>
              <div
                style={{
                  fontSize: 14.5,
                  fontWeight: 800,
                  color: "#ffffff",
                  textTransform: "uppercase",
                  letterSpacing: "0.05em",
                  marginBottom: 8,
                  padding: "3px 8px",
                  borderRadius: 6,
                  display: "inline-block",
                  background: `linear-gradient(135deg, ${C.orange} 0%, #f0964f 100%)`,
                  boxShadow: "0 3px 8px rgba(229,107,45,0.35)",
                }}
              >
                {sk} <span style={{ opacity: 0.85, fontWeight: 700 }}>({parSemestre[sk].length})</span>
              </div>
              <Card style={{ overflow: "hidden" }}>
                <DataTable cols={cols} rows={parSemestre[sk]} emptyMsg="Aucun module trouvé" compact={true} />
              </Card>
            </div>
          ));
        })()}
      {role !== "etudiant" &&
        (() => {
          const filieresOptionsListe = getFilieres(data);
          const semestresOptionsListe = [
            ...new Set((data.presences || []).map((p) => p.semestre).filter(Boolean)),
          ].sort((a, b) => a.localeCompare(b, "fr", { numeric: true, sensitivity: "base" }));
          const modulesOptionsListe = [...new Set((data.presences || []).map((p) => p.module).filter(Boolean))].sort(
            (a, b) => a.localeCompare(b, "fr", { sensitivity: "base" })
          );
          const professeursOptionsListe = [
            ...new Map(
              (data.presences || []).filter((p) => p.enseignantCin).map((p) => [p.enseignantCin, p.enseignant])
            ).entries(),
          ];
          return (
            /* @__PURE__ */ <Card style={{ padding: 16, marginBottom: 16 }}>
              <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(160px, 1fr))", gap: 10 }}>
                <input
                  placeholder="🔍 Nom, N° inscription..."
                  value={filtreRecherche}
                  onChange={(e) => setFiltreRecherche(e.target.value)}
                  style={{
                    padding: "9px 12px",
                    border: `1px solid ${C.border}`,
                    borderRadius: 8,
                    fontSize: 13,
                    fontFamily: "inherit",
                  }}
                />
                <select
                  value={filtreFiliereListe}
                  onChange={(e) => setFiltreFiliereListe(e.target.value)}
                  style={{
                    padding: "9px 12px",
                    border: `1px solid ${C.border}`,
                    borderRadius: 8,
                    fontSize: 13,
                    fontFamily: "inherit",
                    background: C.white,
                  }}
                >
                  <option value="">-- Tous les groupes --</option>
                  {filieresOptionsListe.map((f) => (
                    /* @__PURE__ */ <option key={f} value={f}>
                      {f}
                    </option>
                  ))}
                </select>
                <select
                  value={filtreSemestreListe}
                  onChange={(e) => setFiltreSemestreListe(e.target.value)}
                  style={{
                    padding: "9px 12px",
                    border: `1px solid ${C.border}`,
                    borderRadius: 8,
                    fontSize: 13,
                    fontFamily: "inherit",
                    background: C.white,
                  }}
                >
                  <option value="">-- Tous les semestres --</option>
                  {semestresOptionsListe.map((s) => (
                    /* @__PURE__ */ <option key={s} value={s}>
                      {s}
                    </option>
                  ))}
                </select>
                <select
                  value={filtreModuleListe}
                  onChange={(e) => setFiltreModuleListe(e.target.value)}
                  style={{
                    padding: "9px 12px",
                    border: `1px solid ${C.border}`,
                    borderRadius: 8,
                    fontSize: 13,
                    fontFamily: "inherit",
                    background: C.white,
                  }}
                >
                  <option value="">-- Tous les modules --</option>
                  {modulesOptionsListe.map((m) => (
                    /* @__PURE__ */ <option key={m} value={m}>
                      {m}
                    </option>
                  ))}
                </select>
                <select
                  value={filtreProfesseurListe}
                  onChange={(e) => setFiltreProfesseurListe(e.target.value)}
                  style={{
                    padding: "9px 12px",
                    border: `1px solid ${C.border}`,
                    borderRadius: 8,
                    fontSize: 13,
                    fontFamily: "inherit",
                    background: C.white,
                  }}
                >
                  <option value="">-- Tous les professeurs --</option>
                  {professeursOptionsListe.map(([cin, nom]) => (
                    /* @__PURE__ */ <option key={cin} value={cin}>
                      {nom}
                    </option>
                  ))}
                </select>
                <select
                  value={filtreStatutListe}
                  onChange={(e) => setFiltreStatutListe(e.target.value)}
                  style={{
                    padding: "9px 12px",
                    border: `1px solid ${C.border}`,
                    borderRadius: 8,
                    fontSize: 13,
                    fontFamily: "inherit",
                    background: C.white,
                  }}
                >
                  <option value="">-- Tous les statuts --</option>
                  <option value="Absent">Non justifiées</option>
                  <option value="Justifié">Justifiées</option>
                </select>
                <input
                  type="date"
                  value={filtreDateDebut}
                  onChange={(e) => setFiltreDateDebut(e.target.value)}
                  placeholder="Du"
                  style={{
                    padding: "9px 12px",
                    border: `1px solid ${C.border}`,
                    borderRadius: 8,
                    fontSize: 13,
                    fontFamily: "inherit",
                  }}
                />
                <input
                  type="date"
                  value={filtreDateFin}
                  onChange={(e) => setFiltreDateFin(e.target.value)}
                  placeholder="Au"
                  style={{
                    padding: "9px 12px",
                    border: `1px solid ${C.border}`,
                    borderRadius: 8,
                    fontSize: 13,
                    fontFamily: "inherit",
                  }}
                />
              </div>
            </Card>
          );
        })()}
      {(() => {
        const presencesAffichees =
          role === "etudiant"
            ? mesPresences
            : mesPresences.filter((p) => {
                const matchRecherche =
                  !filtreRecherche ||
                  `${p.etudiant || ""} ${p.numInscription || ""}`.toLowerCase().includes(filtreRecherche.toLowerCase());
                const matchFiliere = !filtreFiliereListe || p.filiere === filtreFiliereListe;
                const matchSemestre = !filtreSemestreListe || p.semestre === filtreSemestreListe;
                const matchModule = !filtreModuleListe || p.module === filtreModuleListe;
                const matchProfesseur = !filtreProfesseurListe || p.enseignantCin === filtreProfesseurListe;
                const matchStatut = !filtreStatutListe || p.statut === filtreStatutListe;
                const matchDebut = !filtreDateDebut || (p.date || "") >= filtreDateDebut;
                const matchFin = !filtreDateFin || (p.date || "") <= filtreDateFin;
                return (
                  matchRecherche &&
                  matchFiliere &&
                  matchSemestre &&
                  matchModule &&
                  matchProfesseur &&
                  matchStatut &&
                  matchDebut &&
                  matchFin
                );
              });
        const handleDeletePresence = (p) => {
          if (!window.confirm(`Supprimer l'absence de "${p.etudiant}" ?`)) return;
          setData((d) => ({ ...d, presences: d.presences.filter((x) => x.id !== p.id) }));
          if (user?.token && p._spId) deleteFromSP("ECOGEST_Presences", p._spId, user.token);
          toast("Absence supprimée de SharePoint ✓");
        };
        const lignesPresences = presencesAffichees
          .map((p) => ({ ...p, _filiereCode: filiereCode(p.filiere, data.filieres) || p.filiere || "" }))
          .sort(
            (a, b) =>
              a._filiereCode.localeCompare(b._filiereCode, "fr", { numeric: true, sensitivity: "base" }) ||
              (a.semestre || "").localeCompare(b.semestre || "", "fr", { numeric: true, sensitivity: "base" }) ||
              (b.date || "").localeCompare(a.date || "")
          );
        const bandeFilierePresParId = {};
        const premiereLigneFilierePresId = {};
        const derniereLigneFilierePresId = {};
        {
          let bande = 0,
            filierePrecedente = null;
          lignesPresences.forEach((e, idx) => {
            if (e._filiereCode !== filierePrecedente) {
              bande = 1 - bande;
              filierePrecedente = e._filiereCode;
              premiereLigneFilierePresId[e.id] = true;
              if (idx > 0) derniereLigneFilierePresId[lignesPresences[idx - 1].id] = true;
            }
            bandeFilierePresParId[e.id] = bande;
          });
          if (lignesPresences.length > 0)
            derniereLigneFilierePresId[lignesPresences[lignesPresences.length - 1].id] = true;
        }
        const cols = [
          { key: "_filiereCode", label: "Filière", render: (v) => /* @__PURE__ */ <FiliereBadge code={v} /> },
          { key: "date", label: "Date", render: (v) => formatDateFR(v) || v },
          {
            key: "etudiant",
            label: "Étudiant",
            render: (v, row) => {
              const justifie =
                row.statut === "Justifié" && (row.documentJustification || row.justifiePar || row.dateJustification);
              const coul = justifie ? C.green : C.red;
              return (
                /* @__PURE__ */ <div>
                  <div>{v}</div>
                  {(justifie || row.statut === "Absent") && (
                    /* @__PURE__ */ <div style={{ marginTop: 3 }}>
                      <span
                        style={{
                          display: "inline-flex",
                          alignItems: "center",
                          gap: 6,
                          background: coul,
                          color: "#ffffff",
                          fontSize: 10.5,
                          fontWeight: 700,
                          padding: "2px 9px 2px 3px",
                          borderRadius: 20,
                          whiteSpace: "nowrap",
                          boxShadow: `0 2px 6px ${coul}70`,
                        }}
                      >
                        <span
                          style={{
                            background: "rgba(255,255,255,0.94)",
                            color: coul,
                            borderRadius: "50%",
                            width: 16,
                            height: 16,
                            display: "flex",
                            alignItems: "center",
                            justifyContent: "center",
                            fontSize: 10,
                            flexShrink: 0,
                          }}
                        >
                          {justifie ? "✓" : "⚠️"}
                        </span>
                        {justifie
                          ? `${row.documentJustification || "Justifié"}${row.justifiePar ? " · par " + row.justifiePar : ""}${row.dateJustification ? " · le " + formatDateFR(row.dateJustification) : ""}`
                          : "Non justifiée"}
                      </span>
                    </div>
                  )}
                </div>
              );
            },
          },
          { key: "semestre", label: "Semestre" },
          { key: "module", label: "Module" },
          { key: "enseignant", label: "Professeur", render: (v) => (role === "etudiant" ? masquerCinProf(v) : v) },
          {
            key: "statut",
            label: "Statut",
            render: (v) => (
              /* @__PURE__ */ <Badge color={v === "Présent" ? C.green : v === "Absent" ? C.red : C.orange}>{v}</Badge>
            ),
          },
          ...(canEdit
            ? [
                {
                  key: "_modifier",
                  label: "Modifier",
                  render: (_, row) => (
                    /* @__PURE__ */ <Btn
                      small={true}
                      variant="navy"
                      onClick={(e) => {
                        e.stopPropagation();
                        openJustif(row);
                      }}
                    >
                      ✏️ Modifier
                    </Btn>
                  ),
                },
              ]
            : []),
          ...(canEdit
            ? [
                {
                  key: "_suppr",
                  label: "Suppr.",
                  render: (_, row) => (
                    /* @__PURE__ */ <Btn
                      small={true}
                      variant="danger"
                      onClick={(e) => {
                        e.stopPropagation();
                        handleDeletePresence(row);
                      }}
                    >
                      🗑
                    </Btn>
                  ),
                },
              ]
            : []),
        ];
        if (lignesPresences.length === 0)
          return (
            /* @__PURE__ */ <Card style={{ padding: 32, textAlign: "center", color: C.muted, fontSize: 13 }}>
              Aucune absence enregistrée
            </Card>
          );
        return (
          /* @__PURE__ */ <Card style={{ overflow: "hidden" }}>
            <DataTable
              cols={cols}
              rows={lignesPresences}
              emptyMsg="Aucune absence trouvée"
              compact={true}
              onRowClick={canEdit ? openJustif : void 0}
              rowStyle={(row) => ({
                background: bandeFilierePresParId[row.id] === 0 ? C.navyLight : "#e0f4f2",
                borderTop: premiereLigneFilierePresId[row.id] ? `2.5px solid ${C.navy}` : void 0,
                borderBottom: derniereLigneFilierePresId[row.id] ? `2.5px solid ${C.navy}` : `1px solid ${C.border}`,
              })}
            />
          </Card>
        );
      })()}
      {showModal && (
        /* @__PURE__ */ <Modal
          title="Enregistrer des absences → SharePoint"
          onClose={() => setShowModal(false)}
          width={640}
        >
          <div className="eco-grid2" style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "0 14px" }}>
            <Input
              label="Date de l'absence *"
              type="date"
              value={form.date}
              min={semestreInfoForm?.dateDebut || void 0}
              max={semestreInfoForm?.dateFin || void 0}
              onChange={(e) => setForm((f) => ({ ...f, date: e.target.value }))}
            />
            <Sel
              label="Professeur * (liste — page Professeurs)"
              value={form.professeurCin}
              onChange={(e) => {
                const p = professeursRecords.find((pr) => pr.cin === e.target.value);
                setForm((f) => ({
                  ...f,
                  professeurCin: e.target.value,
                  professeur: p ? `${p.prenom} ${p.nom}` : "",
                  module: "",
                }));
              }}
            >
              <option value="">
                {professeursRecords.length
                  ? "-- Sélectionner --"
                  : "-- Aucun professeur enregistré (page Professeurs) --"}
              </option>
              {professeursRecords.map((p) => (
                /* @__PURE__ */ <option
                  key={p.cin || p.id}
                  value={p.cin}
                >{`${p.prenom} ${p.nom} — CIN ${p.cin || "non renseigné"}`}</option>
              ))}
            </Sel>
          </div>
          <div className="eco-grid2" style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "0 14px" }}>
            <Sel
              label="Filière *"
              value={form.filiere}
              onChange={(e) => {
                setForm((f) => ({ ...f, filiere: e.target.value, module: "" }));
                setAbsents({});
              }}
            >
              <option value="">-- Sélectionner --</option>
              {filieresOptions.map((f) => (
                /* @__PURE__ */ <option key={f} value={f}>
                  {f}
                </option>
              ))}
            </Sel>
            <Sel
              label="Semestre *"
              value={form.semestre}
              onChange={(e) => {
                const sInfo = (data.semestresDates || []).find((s) => s.semestre === e.target.value);
                setForm((f) => ({ ...f, semestre: e.target.value, module: "", date: sInfo?.dateDebut || f.date }));
              }}
            >
              <option value="">-- Sélectionner --</option>
              {semestresOptions.map((s) => (
                /* @__PURE__ */ <option key={s}>{s}</option>
              ))}
            </Sel>
          </div>
          {semestreInfoForm && (semestreInfoForm.dateDebut || semestreInfoForm.dateFin) && (
            /* @__PURE__ */ <div style={{ fontSize: 11, color: C.muted, marginTop: -10, marginBottom: 13 }}>
              📅 Période du semestre (Paramétrage) : {formatDateFR(semestreInfoForm.dateDebut) || "—"} → {formatDateFR(semestreInfoForm.dateFin) || "—"}
            </div>
          )}
          <div className="eco-grid2" style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "0 14px" }}>
            <Sel
              label="Module *"
              value={form.module}
              onChange={(e) => setForm((f) => ({ ...f, module: e.target.value }))}
            >
              <option value="">-- Sélectionner --</option>
              {modulesOptions.map((c) => (
                /* @__PURE__ */ <option key={c.id} value={c.intitule}>
                  {c.intitule}
                </option>
              ))}
            </Sel>
            <Sel
              label="Séance *"
              value={form.seance}
              onChange={(e) => setForm((f) => ({ ...f, seance: e.target.value }))}
            >
              <option value="">-- Sélectionner --</option>
              {SEANCES_HORAIRE.map((s) => (
                /* @__PURE__ */ <option key={s}>{s}</option>
              ))}
            </Sel>
          </div>
          {form.filiere && form.semestre && modulesOptions.length === 0 && (
            /* @__PURE__ */ <div
              style={{
                background: "#fff8e6",
                border: "1px solid #f0d078",
                borderRadius: 8,
                padding: "9px 13px",
                fontSize: 12,
                color: "#7a5c00",
                marginBottom: 13,
              }}
            >
              ⚠️ Aucun module trouvé pour cette filière et ce semestre dans la page Modules.
            </div>
          )}
          {form.filiere && (
            /* @__PURE__ */ <div style={{ marginTop: 6, marginBottom: 16 }}>
              <div
                style={{
                  fontSize: 11.5,
                  fontWeight: 700,
                  color: C.red,
                  textTransform: "uppercase",
                  letterSpacing: "0.06em",
                  marginBottom: 8,
                }}
              >
                Cochez les étudiants absents ({etudiantsFiliere.length} étudiant(s) dans ce groupe
                {nbCoches > 0 ? ` — ${nbCoches} coché(s)` : ""})
              </div>
              {etudiantsFiliere.length === 0 ? (
                /* @__PURE__ */ <div
                  style={{
                    textAlign: "center",
                    padding: 16,
                    color: C.muted,
                    fontSize: 12.5,
                    border: `1px solid ${C.border}`,
                    borderRadius: 8,
                  }}
                >
                  Aucun étudiant dans ce groupe
                </div>
              ) : (
                /* @__PURE__ */ <div
                  style={{ border: `1px solid ${C.border}`, borderRadius: 10, maxHeight: 260, overflowY: "auto" }}
                >
                  {etudiantsFiliere.map((et, i) => (
                    /* @__PURE__ */ <label
                      key={et.id}
                      style={{
                        display: "flex",
                        alignItems: "center",
                        gap: 10,
                        padding: "9px 14px",
                        borderBottom: i < etudiantsFiliere.length - 1 ? `1px solid ${C.light}` : "none",
                        background: absents[et.id] ? "#fdeeec" : i % 2 === 0 ? C.white : C.light,
                        cursor: "pointer",
                      }}
                    >
                      <input
                        type="checkbox"
                        checked={!!absents[et.id]}
                        onChange={() => toggleAbsent(et.id)}
                        style={{ width: 16, height: 16, cursor: "pointer" }}
                      />
                      <span style={{ fontSize: 13, fontWeight: 600, color: absents[et.id] ? C.red : "#232733" }}>
                        {et.prenom} {et.nom}
                      </span>
                      <span style={{ fontSize: 11, color: C.muted, marginLeft: "auto" }}>{et.numInscription}</span>
                    </label>
                  ))}
                </div>
              )}
            </div>
          )}
          <div style={{ display: "flex", gap: 10, justifyContent: "flex-end" }}>
            <Btn variant="light" onClick={() => setShowModal(false)}>
              Annuler
            </Btn>
            <Btn
              onClick={handleSave}
              disabled={
                !form.date ||
                !form.professeurCin ||
                !form.module ||
                !form.seance ||
                !form.filiere ||
                !form.semestre ||
                nbCoches === 0 ||
                saving
              }
            >
              {saving ? "Enregistrement..." : `Enregistrer ${nbCoches > 0 ? `(${nbCoches})` : ""} → SharePoint`}
            </Btn>
          </div>
        </Modal>
      )}
      {showJustif && (
        /* @__PURE__ */ <Modal
          title={`Justification de l'absence — ${showJustif.etudiant}`}
          onClose={() => setShowJustif(null)}
        >
          <div
            style={{
              background: C.light,
              border: `1px solid ${C.border}`,
              borderRadius: 8,
              padding: "10px 14px",
              fontSize: 12.5,
              color: "#444",
              marginBottom: 16,
              lineHeight: 1.7,
            }}
          >
            <strong>{showJustif.etudiant}</strong> · {showJustif.filiere || "—"}
            <br />
            {formatDateFR(showJustif.date)} · {showJustif.module || "—"} · {showJustif.seance || showJustif.heure || "—"}
          </div>
          <Sel
            label="Statut *"
            value={justifForm.statut}
            onChange={(e) => setJustifForm((f) => ({ ...f, statut: e.target.value }))}
          >
            <option>Absent</option>
            <option>Justifié</option>
          </Sel>
          {justifForm.statut === "Justifié" && (
            /* @__PURE__ */ <React.Fragment>
              <CreatableSelect
                label="Document de justification"
                value={justifForm.documentJustification}
                onChange={(v) => setJustifForm((f) => ({ ...f, documentJustification: v }))}
                options={[
                  "Certificat médical",
                  "Convocation administrative",
                  "Certificat de décès",
                  "Attestation de déplacement",
                ]}
              />
              <div className="eco-grid2" style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "0 14px" }}>
                <CreatableSelect
                  label="Justifié par"
                  value={justifForm.justifiePar}
                  onChange={(v) => setJustifForm((f) => ({ ...f, justifiePar: v }))}
                  options={["Père", "Mère", "Lui même"]}
                />
                <Input
                  label="Date de justification"
                  type="date"
                  value={justifForm.dateJustification}
                  onChange={(e) => setJustifForm((f) => ({ ...f, dateJustification: e.target.value }))}
                />
              </div>
            </React.Fragment>
          )}
          <div style={{ display: "flex", gap: 10, justifyContent: "space-between", alignItems: "center" }}>
            {justifForm.statut === "Justifié" ? (
              /* @__PURE__ */ <Btn
                variant="light"
                onClick={() => genererBilletEntree({ ...showJustif, ...justifForm })}
              >
                🖨️ Imprimer un billet d'entrée
              </Btn>
            ) : (
              /* @__PURE__ */ <span />
            )}
            <div style={{ display: "flex", gap: 10 }}>
              <Btn variant="light" onClick={() => setShowJustif(null)}>
                Annuler
              </Btn>
              <Btn onClick={handleSaveJustif}>Enregistrer → SharePoint</Btn>
            </div>
          </div>
        </Modal>
      )}
    </div>
  );
}
function CreatableSelect({ label, value, onChange, options }) {
  const [adding, setAdding] = useState(false);
  const [newVal, setNewVal] = useState("");
  const confirmAdd = () => {
    const v = newVal.trim();
    if (v) onChange(v);
    setAdding(false);
    setNewVal("");
  };
  if (adding) {
    return (
      /* @__PURE__ */ <div style={{ marginBottom: 13 }}>
        <label style={{ fontSize: 12, fontWeight: 600, color: "#555", display: "block", marginBottom: 4 }}>
          {label}
        </label>
        <div style={{ display: "flex", gap: 6 }}>
          <input
            autoFocus={true}
            value={newVal}
            onChange={(e) => setNewVal(e.target.value)}
            onKeyDown={(e) => {
              if (e.key === "Enter") confirmAdd();
              if (e.key === "Escape") setAdding(false);
            }}
            placeholder="Nouvel élément..."
            style={{
              flex: 1,
              padding: "8px 11px",
              border: `1px solid ${C.navy}`,
              borderRadius: 8,
              fontSize: 13,
              fontFamily: "inherit",
              boxSizing: "border-box",
            }}
          />
          <button
            onClick={confirmAdd}
            title="Ajouter"
            style={{
              border: "none",
              background: C.green,
              color: "white",
              borderRadius: 8,
              padding: "0 13px",
              cursor: "pointer",
              fontWeight: 700,
              fontFamily: "inherit",
            }}
          >
            ✓
          </button>
          <button
            onClick={() => {
              setAdding(false);
              setNewVal("");
            }}
            title="Annuler"
            style={{
              border: "none",
              background: C.light,
              color: "#555",
              borderRadius: 8,
              padding: "0 13px",
              cursor: "pointer",
              fontFamily: "inherit",
            }}
          >
            ✕
          </button>
        </div>
      </div>
    );
  }
  const allOptions = value && !options.includes(value) ? [value, ...options] : options;
  return (
    /* @__PURE__ */ <Sel
      label={label}
      value={value}
      onChange={(e) => {
        if (e.target.value === "__add__") setAdding(true);
        else onChange(e.target.value);
      }}
    >
      <option value="">-- Sélectionner --</option>
      {allOptions.map((o) => (
        /* @__PURE__ */ <option key={o} value={o}>
          {o}
        </option>
      ))}
      <option value="__add__">➕ Ajouter un nouvel élément...</option>
    </Sel>
  );
}
const DEPARTEMENTS_SEED = [
  "Informatique & Systèmes d'Information",
  "Management & Gestion des Entreprises",
  "Finance & Comptabilité",
  "Marketing & Commerce",
  "Langues & Communication",
];
const RESPONSABLES_PEDAGOGIQUES_SEED = ["M. Iliass ELMZIOUI", "Mme Fatima Zahra ZNIBER"];
const SEMESTRES_MODULE_SEED = ["Semestre 1", "Semestre 2"];
const PROFESSEURS_MODULE_SEED = [
  "ABDELHAMID MALTI",
  "ABDELILAH ELHAMYANI",
  "ABDELJEBBAR SALIM",
  "ADILA BAHALLA",
  "CHARAF BRITEL",
  "HICHAM MOHMOH",
  "ILIASS ELMZIOUI",
  "MAJDA ELHOZMARI",
  "MARIEM BEGGAR",
  "NABILA BENOHOUD",
  "SOUKAINA SABIR",
  "YASSIR NAJI",
  "YOUSSEF ZAHIR",
  "KHAWLA MSSYAH",
  "OUSSAMA RAHMOUNI",
  "KHALIL RHATAY",
  "MOHAMED AMINE ESSFALI",
  "NADIA QOUDHADH",
  "KAWTAR OUBDI",
  "ABDELKRIM CHIRIG",
  "HALIMA HACHEM",
  "BADR FAWZI",
  "ADIL KADIRI",
];
function Cours({ data, setData, role, user, toast }) {
  const canEdit = ["administrateur", "directrice"].includes(role);
  const [showModal, setShowModal] = useState(false);
  const [editItem, setEditItem] = useState(null);
  const [filterFiliere, setFilterFiliere] = useState("");
  const [filterSemestre, setFilterSemestre] = useState("");
  const [search, setSearch] = useState("");
  const emptyForm = {
    intitule: "",
    departement: "",
    responsablePedagogique: "",
    semestre: "",
    filieres: [],
    professeur: "",
    professeurCin: "",
    heures: "",
    salle: "",
  };
  const [form, setForm] = useState(emptyForm);
  const optionsFor = (field, seed) => {
    const used = [...new Set((data.cours || []).map((c) => c[field]).filter(Boolean))];
    return [.../* @__PURE__ */ new Set([...seed, ...used])].sort((a, b) =>
      a.localeCompare(b, "fr", { sensitivity: "base" })
    );
  };
  const departements = optionsFor("departement", DEPARTEMENTS_SEED);
  const responsables = optionsFor("responsablePedagogique", RESPONSABLES_PEDAGOGIQUES_SEED);
  const semestres = optionsFor("semestre", SEMESTRES_MODULE_SEED);
  const professeursRecords = [...(data.professeurs || [])].sort((a, b) =>
    `${a.prenom || ""} ${a.nom || ""}`
      .trim()
      .localeCompare(`${b.prenom || ""} ${b.nom || ""}`.trim(), "fr", { sensitivity: "base" })
  );
  const filieresOptions = getFilieres(data);
  const openAdd = () => {
    setEditItem(null);
    setForm(emptyForm);
    setShowModal(true);
  };
  const openEdit = (c) => {
    setEditItem(c);
    setForm({
      ...c,
      heures: String(c.heures || ""),
      filieres: c.filieresListe || (c.filiere ? [c.filiere] : []),
      professeurCin: c.professeurCin || "",
    });
    setShowModal(true);
  };
  const spFields = (f) => ({
    Title: f.intitule || "",
    Intitule: f.intitule || "",
    Departement: f.departement || "",
    ResponsablePedagogique: f.responsablePedagogique || "",
    Semestre: f.semestre || "",
    Filiere: (f.filieres || []).join(" | "),
    Professeur: f.professeur || "",
    ProfesseurCIN: f.professeurCin || "",
    HeuresTotal: String(parseInt(f.heures) || 0),
    Salle: f.salle || "",
  });
  const handleSave = () => {
    if (!user?.token) {
      toast("⚠️ Non connecté à SharePoint — module non enregistré");
      setShowModal(false);
      return;
    }
    const filiereJointe = (form.filieres || []).join(" | ");
    if (editItem) {
      const updated = {
        ...form,
        id: editItem.id,
        _spId: editItem._spId,
        heures: parseInt(form.heures) || 0,
        filiere: filiereJointe,
        filieresListe: form.filieres || [],
      };
      setData((d) => ({ ...d, cours: d.cours.map((c) => (c.id === editItem.id ? updated : c)) }));
      saveToSP("ECOGEST_Cours", spFields(form), user.token, editItem._spId).then((result) => {
        if (!result) {
          toast("⚠️ Échec de la mise à jour dans SharePoint");
          return;
        }
        toast("Module mis à jour dans SharePoint ✓");
      });
    } else {
      const newC = {
        ...form,
        id: Date.now(),
        heures: parseInt(form.heures) || 0,
        filiere: filiereJointe,
        filieresListe: form.filieres || [],
      };
      saveToSP("ECOGEST_Cours", spFields(form), user.token).then((result) => {
        if (!result) {
          toast("⚠️ Échec de l'enregistrement dans SharePoint — module non ajouté");
          return;
        }
        setData((d) => ({ ...d, cours: [...d.cours, { ...newC, _spId: result.id }] }));
        toast("Module enregistré dans SharePoint ✓");
      });
    }
    setShowModal(false);
    setForm(emptyForm);
  };
  const handleDelete = (c) => {
    const presencesLiees = (data.presences || []).filter((p) => p.module === c.intitule);
    const notesLiees = (data.notes || []).filter((n) => n.module === c.intitule);
    const avertissement =
      presencesLiees.length || notesLiees.length
        ? `

⚠️ ${presencesLiees.length} absence(s) et ${notesLiees.length} évaluation(s) liée(s) à ce module seront également supprimées.`
        : "";
    if (!window.confirm(`Supprimer le module "${c.intitule}" ?${avertissement}`)) return;
    setData((d) => ({
      ...d,
      cours: d.cours.filter((x) => x.id !== c.id),
      presences: d.presences.filter((p) => p.module !== c.intitule),
      notes: d.notes.filter((n) => n.module !== c.intitule),
    }));
    if (user?.token) {
      if (c._spId) deleteFromSP("ECOGEST_Cours", c._spId, user.token);
      presencesLiees.forEach((p) => {
        if (p._spId) deleteFromSP("ECOGEST_Presences", p._spId, user.token);
      });
      notesLiees.forEach((n) => {
        if (n._spId) deleteFromSP("ECOGEST_Notes", n._spId, user.token);
      });
    }
    const suffixe =
      presencesLiees.length || notesLiees.length
        ? ` (+ ${presencesLiees.length} absence(s), ${notesLiees.length} évaluation(s))`
        : "";
    toast(`Module supprimé de SharePoint ✓${suffixe}`);
  };
  const groupes = {};
  (data.cours || [])
    .filter((c) => !filterSemestre || c.semestre === filterSemestre)
    .filter(
      (c) =>
        !search ||
        `${c.intitule || ""} ${c.professeur || ""} ${c.departement || ""} ${c.responsablePedagogique || ""}`
          .toLowerCase()
          .includes(search.toLowerCase())
    )
    .forEach((c) => {
      const filieresDuModule =
        c.filieresListe && c.filieresListe.length ? c.filieresListe : [c.filiere || "Sans filière"];
      filieresDuModule.forEach((filiereKey) => {
        if (filterFiliere && filiereKey !== filterFiliere) return;
        const semestreKey = c.semestre || "Sans semestre";
        if (!groupes[filiereKey]) groupes[filiereKey] = {};
        if (!groupes[filiereKey][semestreKey]) groupes[filiereKey][semestreKey] = [];
        groupes[filiereKey][semestreKey].push(c);
      });
    });
  const groupeKeys = Object.keys(groupes).sort((a, b) =>
    a.localeCompare(b, "fr", { numeric: true, sensitivity: "base" })
  );
  const cols = [
    {
      key: "intitule",
      label: "Nom du module",
      render: (v, row) => (
        /* @__PURE__ */ <div>
          <div>{v}</div>
          {(row.filieresListe || []).length > 1 && (
            /* @__PURE__ */ <div style={{ fontSize: 10, color: C.purple, fontWeight: 700, marginTop: 2 }}>
              🔗 Module commun ({row.filieresListe.length} filières)
            </div>
          )}
        </div>
      ),
    },
    { key: "departement", label: "Département" },
    { key: "responsablePedagogique", label: "Responsable pédagogique" },
    { key: "professeur", label: "Professeur" },
    { key: "heures", label: "Volume horaire", render: (v) => `${v}h` },
    ...(canEdit
      ? [
          {
            key: "_modifier",
            label: "Modifier",
            render: (_, row) => (
              /* @__PURE__ */ <Btn small={true} variant="navy" onClick={() => openEdit(row)}>
                ✏️ Modifier
              </Btn>
            ),
          },
        ]
      : []),
    ...(canEdit
      ? [
          {
            key: "_suppr",
            label: "Suppr.",
            render: (_, row) => (
              /* @__PURE__ */ <Btn small={true} variant="danger" onClick={() => handleDelete(row)}>
                🗑
              </Btn>
            ),
          },
        ]
      : []),
  ];
  return (
    /* @__PURE__ */ <div>
      <PageHeader
        title="Modules"
        sub={`${data.cours.length} module(s) — liste SharePoint "${SP.lists.cours}"`}
        action={canEdit && /* @__PURE__ */ <Btn onClick={openAdd}>+ Ajouter un module</Btn>}
      />
      <div style={{ marginBottom: 16, display: "flex", gap: 10, flexWrap: "wrap" }}>
        <input
          placeholder="🔍  Nom du module, professeur, département..."
          value={search}
          onChange={(e) => setSearch(e.target.value)}
          style={{
            flex: 1,
            minWidth: 220,
            padding: "9px 14px",
            border: `1px solid ${C.border}`,
            borderRadius: 8,
            fontSize: 13,
            boxSizing: "border-box",
            fontFamily: "inherit",
          }}
        />
        <select
          value={filterFiliere}
          onChange={(e) => setFilterFiliere(e.target.value)}
          style={{
            padding: "9px 14px",
            border: `1px solid ${C.border}`,
            borderRadius: 8,
            fontSize: 13,
            fontFamily: "inherit",
            background: C.white,
            minWidth: 260,
          }}
        >
          <option value="">-- Tous les groupes --</option>
          {filieresOptions.map((f) => (
            /* @__PURE__ */ <option key={f} value={f}>
              {f}
            </option>
          ))}
        </select>
        <select
          value={filterSemestre}
          onChange={(e) => setFilterSemestre(e.target.value)}
          style={{
            padding: "9px 14px",
            border: `1px solid ${C.border}`,
            borderRadius: 8,
            fontSize: 13,
            fontFamily: "inherit",
            background: C.white,
            minWidth: 180,
          }}
        >
          <option value="">-- Tous les semestres --</option>
          {semestres.map((s) => (
            /* @__PURE__ */ <option key={s} value={s}>
              {s}
            </option>
          ))}
        </select>
      </div>
      {groupeKeys.length === 0 ? (
        /* @__PURE__ */ <Card style={{ padding: 32, textAlign: "center", color: C.muted, fontSize: 13 }}>
          Aucun module trouvé
        </Card>
      ) : (
        groupeKeys.map((fk) => {
          const semestreGroupes = groupes[fk];
          const semestreKeys = Object.keys(semestreGroupes).sort((a, b) =>
            a.localeCompare(b, "fr", { numeric: true, sensitivity: "base" })
          );
          return (
            /* @__PURE__ */ <div key={fk} style={{ marginBottom: 28 }}>
              <div
                style={{
                  fontSize: 14.5,
                  fontWeight: 800,
                  color: "#ffffff",
                  textTransform: "uppercase",
                  letterSpacing: "0.05em",
                  marginBottom: 12,
                  padding: "3px 8px",
                  borderRadius: 6,
                  display: "inline-block",
                  background: `linear-gradient(135deg, ${C.orange} 0%, #f0964f 100%)`,
                  boxShadow: "0 3px 8px rgba(229,107,45,0.35)",
                }}
              >
                {fk} <span style={{ opacity: 0.85, fontWeight: 700 }}>
                  ({semestreKeys.reduce((s, sk) => s + semestreGroupes[sk].length, 0)})
                </span>
              </div>
              {semestreKeys.map((sk) => (
                /* @__PURE__ */ <div key={sk} style={{ marginBottom: 16, marginLeft: 8 }}>
                  <div style={{ fontSize: 12, fontWeight: 700, color: "#555", marginBottom: 8 }}>
                    {sk} <span style={{ color: C.muted, fontWeight: 600 }}>({semestreGroupes[sk].length})</span>
                  </div>
                  <Card style={{ overflow: "hidden" }}>
                    <DataTable cols={cols} rows={semestreGroupes[sk]} emptyMsg="Aucun module trouvé" />
                  </Card>
                </div>
              ))}
            </div>
          );
        })
      )}
      {showModal && (
        /* @__PURE__ */ <Modal
          title={editItem ? "Modifier le module" : "Ajouter un module → SharePoint"}
          onClose={() => setShowModal(false)}
        >
          <Input
            label="Nom du module *"
            value={form.intitule}
            onChange={(e) => setForm((f) => ({ ...f, intitule: e.target.value }))}
          />
          <div className="eco-grid2" style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "0 14px" }}>
            <CreatableSelect
              label="Département"
              value={form.departement}
              onChange={(v) => setForm((f) => ({ ...f, departement: v }))}
              options={departements}
            />
            <CreatableSelect
              label="Responsable pédagogique"
              value={form.responsablePedagogique}
              onChange={(v) => setForm((f) => ({ ...f, responsablePedagogique: v }))}
              options={responsables}
            />
          </div>
          <CreatableSelect
            label="Semestre"
            value={form.semestre}
            onChange={(v) => setForm((f) => ({ ...f, semestre: v }))}
            options={semestres}
          />
          <div style={{ marginBottom: 13 }}>
            <label style={{ fontSize: 12, fontWeight: 600, color: "#555", display: "block", marginBottom: 6 }}>
              Filière(s) * <span style={{ fontWeight: 400, color: C.muted }}>— cochez plusieurs filières pour un module commun</span>
            </label>
            <div style={{ border: `1px solid ${C.border}`, borderRadius: 8, maxHeight: 160, overflowY: "auto" }}>
              {filieresOptions.map((f, i) => (
                /* @__PURE__ */ <label
                  key={f}
                  style={{
                    display: "flex",
                    alignItems: "center",
                    gap: 9,
                    padding: "8px 12px",
                    borderBottom: i < filieresOptions.length - 1 ? `1px solid ${C.light}` : "none",
                    cursor: "pointer",
                    fontSize: 12.5,
                  }}
                >
                  <input
                    type="checkbox"
                    checked={(form.filieres || []).includes(f)}
                    onChange={() =>
                      setForm((fm) => ({
                        ...fm,
                        filieres: (fm.filieres || []).includes(f)
                          ? fm.filieres.filter((x) => x !== f)
                          : [...(fm.filieres || []), f],
                      }))
                    }
                    style={{ width: 15, height: 15, cursor: "pointer" }}
                  />
                  {f}
                </label>
              ))}
            </div>
          </div>
          <div className="eco-grid2" style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "0 14px" }}>
            <Sel
              label="Professeur * (liste — page Professeurs)"
              value={form.professeurCin}
              onChange={(e) => {
                const p = professeursRecords.find((pr) => pr.cin === e.target.value);
                setForm((f) => ({ ...f, professeurCin: e.target.value, professeur: p ? `${p.prenom} ${p.nom}` : "" }));
              }}
            >
              <option value="">
                {professeursRecords.length
                  ? "-- Sélectionner --"
                  : "-- Aucun professeur enregistré (page Professeurs) --"}
              </option>
              {professeursRecords.map((p) => (
                /* @__PURE__ */ <option
                  key={p.cin || p.id}
                  value={p.cin}
                >{`${p.prenom} ${p.nom} — CIN ${p.cin || "non renseigné"}`}</option>
              ))}
            </Sel>
            <Input
              label="Volume horaire (h)"
              type="number"
              value={form.heures}
              onChange={(e) => setForm((f) => ({ ...f, heures: e.target.value }))}
            />
          </div>
          <div style={{ display: "flex", gap: 10, justifyContent: "flex-end" }}>
            <Btn variant="light" onClick={() => setShowModal(false)}>
              Annuler
            </Btn>
            <Btn onClick={handleSave} disabled={!form.intitule || !form.professeurCin || !(form.filieres || []).length}>
              {editItem ? "Mettre à jour → SharePoint" : "Enregistrer → SharePoint"}
            </Btn>
          </div>
        </Modal>
      )}
    </div>
  );
}
const TYPES_DOCUMENT_DEMANDE = [
  "Attestation d'inscription",
  "Attestation de scolarité",
  "Attestation de réussite",
  "Relevé de notes",
  "Autre",
];
const STATUTS_DEMANDE = ["En attente", "En cours", "Prête", "Délivrée", "Rejetée"];
const STATUT_COLOR = {
  "En attente": "#c99a2e",
  "En cours": "#0d9488",
  Prête: "#2e7d52",
  Délivrée: "#1a2f5e",
  Rejetée: "#c0392b",
};
const MOIS_ANNEE = [
  "janvier",
  "février",
  "mars",
  "avril",
  "mai",
  "juin",
  "juillet",
  "août",
  "septembre",
  "octobre",
  "novembre",
  "décembre",
];
const EMAIL_VALIDATEUR_SEANCES = "i.elmzioui@esrmi.ma";
const EMAILS_ACCUSE_RECEPTION = ["k.bouabid@esrmi.ma"];
const EMAILS_TRAITEMENT_PAIEMENT = ["fz.zniber@esrmi.ma"];
function Pointage({ data, setData, role, user, toast, pointageFiltreCible, setPointageFiltreCible }) {
  const estProf = role === "professeur";
  const estValidateur = ["administrateur", "directrice"].includes(role);
  const monEmail = (user?.email || "").toLowerCase();
  const estIliass = monEmail === EMAIL_VALIDATEUR_SEANCES;
  const estKawtar = EMAILS_ACCUSE_RECEPTION.includes(monEmail);
  const estFatimaZahra = EMAILS_TRAITEMENT_PAIEMENT.includes(monEmail);
  const monProfesseur = (data.professeurs || []).find(
    (p) => p.email && user?.email && p.email.toLowerCase() === user.email.toLowerCase()
  );
  const [showFicheId, setShowFicheId] = useState(null);
  const [showNouvelleFiche, setShowNouvelleFiche] = useState(false);
  const [moisNouvelleFiche, setMoisNouvelleFiche] = useState(/* @__PURE__ */ new Date().toISOString().slice(0, 7));
  const [formSeance, setFormSeance] = useState({
    date: today(),
    module: "",
    heureDebut: "08:00",
    heureFin: "10:00",
    commentaire: "",
  });
  const [editingSeanceId, setEditingSeanceId] = useState(null);
  const [filtreValidateur, setFiltreValidateur] = useState("Soumise");
  useEffect(() => {
    if (pointageFiltreCible) {
      setFiltreValidateur(pointageFiltreCible);
      setShowFicheId(null);
      setPointageFiltreCible?.(null);
    }
  }, [pointageFiltreCible]); // Le module Pointage doit toujours refl\xE9ter les vraies donn\xE9es SharePoint
  // pour tout le monde (professeur, Iliass, Kawtar, Yasser, Fatima Zahra,
  // Abdelilah) \u2014 jamais une copie locale p\xE9rim\xE9e propre \xE0 une session :
  // 1) rafra\xEEchissement d\xE8s l'arriv\xE9e sur la page (liste ou fiche),
  // 2) rafra\xEEchissement suppl\xE9mentaire \xE0 l'ouverture d'une fiche pr\xE9cise,
  // 3) rafra\xEEchissement automatique toutes les 30 secondes tant que la page
  //    Pointage reste ouverte, pour que toutes les sessions convergent vite
  //    vers le m\xEAme \xE9tat sans action manuelle.
  useEffect(() => {
    if (!user?.token) return;
    const rafraichir = () =>
      loadAllFromSP(user.token)
        .then((spData) => {
          if (spData) setData(spData);
        })
        .catch((err) => console.warn("Rafraîchissement Pointage échoué:", err.message));
    rafraichir();
    const intervalle = setInterval(rafraichir, 3e4);
    return () => clearInterval(intervalle);
  }, []);
  useEffect(() => {
    if (showFicheId && user?.token) {
      loadAllFromSP(user.token)
        .then((spData) => {
          if (spData) setData(spData);
        })
        .catch((err) => console.warn("Rafraîchissement de la fiche échoué:", err.message));
    }
  }, [showFicheId]);
  const handleRafraichirManuel = () => {
    if (!user?.token) return;
    loadAllFromSP(user.token)
      .then((spData) => {
        if (spData) {
          setData(spData);
          toast("Données actualisées ✓");
        }
      })
      .catch((err) => console.warn("Rafraîchissement échoué:", err.message));
  };
  const formatMoisLabel = (m) => {
    if (!m) return "";
    const [y, mo] = m.split("-");
    const nom = MOIS_ANNEE[parseInt(mo, 10) - 1] || "";
    return `${nom.charAt(0).toUpperCase()}${nom.slice(1)} ${y}`;
  }; // Une fiche "du mois M" correspond \xE0 la p\xE9riode allant du 26 du mois
  // pr\xE9c\xE9dent au 25 du mois M (ex: fiche "Ao\xFBt 2026" = 26/07/2026 \u2192
  // 25/08/2026). Cette r\xE8gle s'applique \xE0 tous les mois, sans exception.
  const periodeDuMois = (moisStr) => {
    if (!moisStr) return null;
    const [y, m] = moisStr.split("-").map(Number);
    if (!y || !m) return null;
    const fmt = (d) =>
      `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
    return { debut: fmt(new Date(y, m - 2, 26)), fin: fmt(new Date(y, m - 1, 25)) };
  };
  const seancesDeFiche = (ficheId) => (data.pointageSeances || []).filter((s) => s.ficheId === ficheId);
  const fichesDeCeProf =
    estProf && monProfesseur
      ? (data.pointageFiches || [])
          .filter((f) => f.professeurId === monProfesseur.id)
          .sort((a, b) => b.mois.localeCompare(a.mois))
      : [];
  const handleCreerFiche = () => {
    if (!monProfesseur) {
      toast("⚠️ Aucune fiche professeur trouvée pour votre compte — contactez l'administrateur.");
      return;
    }
    if (fichesDeCeProf.some((f) => f.mois === moisNouvelleFiche)) {
      toast("⚠️ Une fiche existe déjà pour ce mois");
      return;
    }
    const nouvelId = Date.now() + Math.random();
    const professeurNom = `${monProfesseur.prenom} ${monProfesseur.nom}`;
    const nouvelleFiche = {
      id: nouvelId,
      _spId: null,
      professeurId: monProfesseur.id,
      professeurNom,
      mois: moisNouvelleFiche,
      statut: "Brouillon",
      dateCreation: today(),
      dateSoumission: "",
      dateValidation: "",
      valideePar: "",
      dateAccuseReception: "",
      accuseReceptionPar: "",
      dateTraitement: "",
      traitePar: "",
    };
    setData((d) => ({ ...d, pointageFiches: [...(d.pointageFiches || []), nouvelleFiche] }));
    if (user?.token) {
      saveToSP(
        "ECOGEST_PointageFiches",
        {
          Title: `${professeurNom} - ${moisNouvelleFiche}`,
          ProfesseurId: String(monProfesseur.id),
          ProfesseurNom: professeurNom,
          Mois: moisNouvelleFiche,
          Statut: "Brouillon",
          DateCreation: today(),
        },
        user.token,
        null
      ).then((result) => {
        if (result)
          setData((d) => ({
            ...d,
            pointageFiches: d.pointageFiches.map((f) => (f.id === nouvelId ? { ...f, _spId: result.id } : f)),
          }));
      });
    }
    setShowNouvelleFiche(false);
    setShowFicheId(nouvelId);
    toast("Fiche de pointage créée ✓");
  };
  const openEditSeance = (seance) => {
    setEditingSeanceId(seance.id);
    setFormSeance({
      date: seance.date || today(),
      module: seance.module || "",
      heureDebut: seance.heureDebut || "08:00",
      heureFin: seance.heureFin || "10:00",
      commentaire: seance.commentaire || "",
    });
  };
  const annulerEditionSeance = () => {
    setEditingSeanceId(null);
    setFormSeance({ date: today(), module: "", heureDebut: "08:00", heureFin: "10:00", commentaire: "" });
  };
  const handleEnregistrerSeance = (fiche) => {
    if (!formSeance.module || !formSeance.date || !formSeance.heureDebut || !formSeance.heureFin) {
      toast("⚠️ Remplissez tous les champs");
      return;
    }
    const periode = periodeDuMois(fiche.mois);
    if (periode && (formSeance.date < periode.debut || formSeance.date > periode.fin)) {
      toast(
        `⚠️ La date doit être comprise entre le ${formatDateFR(periode.debut)} et le ${formatDateFR(periode.fin)} (période de la fiche ${formatMoisLabel(fiche.mois)})`
      );
      return;
    }
    const cours = (data.cours || []).find((c) => c.intitule === formSeance.module);
    const [h1, m1] = formSeance.heureDebut.split(":").map(Number);
    const [h2, m2] = formSeance.heureFin.split(":").map(Number);
    const nbHeures = Math.round(Math.max(0, ((h2 * 60 + m2 - (h1 * 60 + m1)) / 60) * 100)) / 100;
    if (nbHeures <= 0) {
      toast("⚠️ L'heure de fin doit être après l'heure de début");
      return;
    }
    const champsCommuns = {
      date: formSeance.date,
      module: formSeance.module,
      filiere: cours?.filiere || "",
      semestre: cours?.semestre || "",
      heureDebut: formSeance.heureDebut,
      heureFin: formSeance.heureFin,
      nbHeures,
      commentaire: formSeance.commentaire || "",
    };
    const champsSP = {
      DatePointage: formSeance.date,
      Module: formSeance.module,
      Filiere: cours?.filiere || "",
      Semestre: cours?.semestre || "",
      HeureDebut: formSeance.heureDebut,
      HeureFin: formSeance.heureFin,
      NbHeures: String(nbHeures),
      Commentaire: formSeance.commentaire || "",
    };
    if (editingSeanceId) {
      const seanceExistante = (data.pointageSeances || []).find((s) => s.id === editingSeanceId); // Toute modification d'une s\xE9ance annule sa confirmation \u2014 le
      // professeur doit la reconfirmer avant de pouvoir soumettre la fiche.
      setData((d) => ({
        ...d,
        pointageSeances: d.pointageSeances.map((s) =>
          s.id === editingSeanceId ? { ...s, ...champsCommuns, confirmee: false } : s
        ),
      }));
      if (user?.token && seanceExistante?._spId)
        saveToSP("ECOGEST_PointageSeances", { ...champsSP, Confirmee: "false" }, user.token, seanceExistante._spId);
      toast("Séance modifiée — à reconfirmer ✓");
      annulerEditionSeance();
      return;
    }
    const nouvelId = Date.now() + Math.random();
    const nouvelleSeance = {
      id: nouvelId,
      _spId: null,
      ficheId: fiche.id,
      professeurId: fiche.professeurId,
      professeurNom: fiche.professeurNom,
      statut: "EnAttente",
      confirmee: false,
      valideePar: "",
      dateValidation: "",
      ...champsCommuns,
    };
    setData((d) => ({ ...d, pointageSeances: [...(d.pointageSeances || []), nouvelleSeance] }));
    if (user?.token) {
      saveToSP(
        "ECOGEST_PointageSeances",
        {
          Title: `${fiche.professeurNom} - ${formSeance.date}`,
          FicheId: String(fiche.id),
          ProfesseurId: String(fiche.professeurId),
          ProfesseurNom: fiche.professeurNom,
          Statut: "EnAttente",
          Confirmee: "false",
          ...champsSP,
        },
        user.token,
        null
      ).then((result) => {
        if (result)
          setData((d) => ({
            ...d,
            pointageSeances: d.pointageSeances.map((s) => (s.id === nouvelId ? { ...s, _spId: result.id } : s)),
          }));
      });
    }
    setFormSeance({ date: today(), module: "", heureDebut: "08:00", heureFin: "10:00", commentaire: "" });
    toast("Séance ajoutée – à confirmer avant soumission ✓");
  };
  const handleSupprimerSeance = (seance) => {
    if (!window.confirm("Supprimer cette séance ?")) return;
    setData((d) => ({ ...d, pointageSeances: d.pointageSeances.filter((s) => s.id !== seance.id) }));
    if (user?.token && seance._spId) deleteFromSP("ECOGEST_PointageSeances", seance._spId, user.token);
    if (editingSeanceId === seance.id) annulerEditionSeance();
  };
  const handleConfirmerSeance = (seance) => {
    setData((d) => ({
      ...d,
      pointageSeances: d.pointageSeances.map((s) => (s.id === seance.id ? { ...s, confirmee: true } : s)),
    }));
    if (user?.token && seance._spId)
      saveToSP("ECOGEST_PointageSeances", { Confirmee: "true" }, user.token, seance._spId);
  };
  const handleSupprimerFiche = (fiche) => {
    if (role !== "administrateur" && !(role === "professeur" && fiche.statut === "Brouillon")) {
      toast("⚠️ Vous n'avez pas les droits pour supprimer cette fiche de pointage");
      return;
    }
    const seancesLiees = seancesDeFiche(fiche.id);
    const avertissementEtape =
      fiche.statut !== "Brouillon"
        ? ` Cette fiche a déjà progressé dans le circuit de validation (statut actuel : ${badgeStatutFicheTexte(fiche.statut)}) — la supprimer effacera aussi cet historique.`
        : "";
    if (
      !window.confirm(
        `Supprimer définitivement cette fiche de pointage (${fiche.professeurNom} — ${formatMoisLabel(fiche.mois)}) et ses ${seancesLiees.length} séance(s) ?${avertissementEtape} Cette action est irréversible.`
      )
    )
      return;
    setData((d) => ({
      ...d,
      pointageFiches: d.pointageFiches.filter((f) => f.id !== fiche.id),
      pointageSeances: d.pointageSeances.filter((s) => s.ficheId !== fiche.id),
    }));
    if (user?.token) {
      if (fiche._spId) deleteFromSP("ECOGEST_PointageFiches", fiche._spId, user.token);
      seancesLiees.forEach((s) => {
        if (s._spId) deleteFromSP("ECOGEST_PointageSeances", s._spId, user.token);
      });
    }
    if (showFicheId === fiche.id) setShowFicheId(null);
    toast("Fiche de pointage supprimée ✓");
  };
  const handleSoumettreFiche = (fiche) => {
    const seances = seancesDeFiche(fiche.id);
    if (seances.length === 0) {
      toast("⚠️ Ajoutez au moins une séance avant de soumettre");
      return;
    }
    if (seances.some((s) => !s.confirmee)) {
      toast("⚠️ Confirmez toutes les séances avant de soumettre la fiche");
      return;
    }
    if (
      !window.confirm(
        `Soumettre cette fiche de pointage (${seances.length} séance(s)) pour validation ? Vous ne pourrez plus la modifier après soumission.`
      )
    )
      return;
    const maj = { ...fiche, statut: "Soumise", dateSoumission: today() };
    setData((d) => ({ ...d, pointageFiches: d.pointageFiches.map((f) => (f.id === fiche.id ? maj : f)) }));
    if (user?.token && fiche._spId)
      saveToSP("ECOGEST_PointageFiches", { Statut: "Soumise", DateSoumission: today() }, user.token, fiche._spId);
    toast("Fiche soumise pour validation ✓");
  };
  const handleValiderSeance = (seance, statut) => {
    const maj = { ...seance, statut, valideePar: user?.nom || user?.email || "", dateValidation: today() };
    setData((d) => ({ ...d, pointageSeances: d.pointageSeances.map((s) => (s.id === seance.id ? maj : s)) }));
    if (user?.token && seance._spId)
      saveToSP(
        "ECOGEST_PointageSeances",
        { Statut: statut, ValideePar: user?.nom || user?.email || "", DateValidation: today() },
        user.token,
        seance._spId
      );
  };
  const handleCommentaireAdminSeance = (seance, commentaire) => {
    setData((d) => ({
      ...d,
      pointageSeances: d.pointageSeances.map((s) => (s.id === seance.id ? { ...s, commentaire } : s)),
    }));
    if (user?.token && seance._spId)
      saveToSP("ECOGEST_PointageSeances", { Commentaire: commentaire }, user.token, seance._spId);
  }; // \xC9tape 1 (M. Iliass Elmzioui uniquement) : une fois toutes les s\xE9ances
  // valid\xE9es ou rejet\xE9es, la fiche passe \xE0 "ValideeIliass" et est
  // transmise \xE0 Mme Bouabid pour accus\xE9 de r\xE9ception.
  const handleValiderFiche = (fiche) => {
    if (
      !window.confirm(
        "Marquer les séances de cette fiche comme validées et la transmettre à Mme Kawtar Bouabid pour accusé de réception ?"
      )
    )
      return;
    const maj = {
      ...fiche,
      statut: "ValideeIliass",
      dateValidation: today(),
      valideePar: user?.nom || user?.email || "",
    };
    setData((d) => ({ ...d, pointageFiches: d.pointageFiches.map((f) => (f.id === fiche.id ? maj : f)) }));
    if (user?.token && fiche._spId)
      saveToSP(
        "ECOGEST_PointageFiches",
        { Statut: "ValideeIliass", DateValidation: today(), ValideePar: user?.nom || user?.email || "" },
        user.token,
        fiche._spId
      );
    toast("Fiche validée — transmise à Mme Bouabid ✓");
  }; // \xC9tape 2 (Mme Kawtar Bouabid uniquement) : accus\xE9 de r\xE9ception de la
  // fiche, transmission \xE0 Mme Zniber pour traitement du paiement.
  const handleAccuserReception = (fiche) => {
    if (
      !window.confirm(
        "Accuser réception de cette fiche et la transmettre à Mme Fatima Zahra Zniber pour traitement du paiement ?"
      )
    )
      return;
    const maj = {
      ...fiche,
      statut: "ReceptionAccusee",
      dateAccuseReception: today(),
      accuseReceptionPar: user?.nom || user?.email || "",
    };
    setData((d) => ({ ...d, pointageFiches: d.pointageFiches.map((f) => (f.id === fiche.id ? maj : f)) }));
    if (user?.token && fiche._spId)
      saveToSP(
        "ECOGEST_PointageFiches",
        {
          Statut: "ReceptionAccusee",
          DateAccuseReception: today(),
          AccuseReceptionPar: user?.nom || user?.email || "",
        },
        user.token,
        fiche._spId
      );
    toast("Réception accusée — transmise à Mme Zniber ✓");
  }; // \xC9tape 3 (Mme Fatima Zahra Zniber uniquement) : confirmation du
  // traitement du paiement, la fiche est marqu\xE9e enti\xE8rement termin\xE9e.
  const handleConfirmerTraitement = (fiche) => {
    if (!window.confirm("Confirmer que le paiement de cette fiche a été traité, et marquer la fiche comme terminée ?"))
      return;
    const maj = { ...fiche, statut: "Terminee", dateTraitement: today(), traitePar: user?.nom || user?.email || "" };
    setData((d) => ({ ...d, pointageFiches: d.pointageFiches.map((f) => (f.id === fiche.id ? maj : f)) }));
    if (user?.token && fiche._spId)
      saveToSP(
        "ECOGEST_PointageFiches",
        { Statut: "Terminee", DateTraitement: today(), TraitePar: user?.nom || user?.email || "" },
        user.token,
        fiche._spId
      );
    toast("Paiement traité — fiche terminée ✓");
  };
  const MAP_STATUTS_FICHE = {
    Brouillon: [C.muted, "📝 Brouillon"],
    Soumise: [C.orange, "⏳ À valider (Iliass)"],
    ValideeIliass: [C.teal, "📤 À réceptionner"],
    ReceptionAccusee: [C.purple, "💰 À traiter"],
    Terminee: [C.green, "✅ Terminée"],
  };
  const badgeStatutFiche = (statut) => {
    const [color, label] = MAP_STATUTS_FICHE[statut] || [C.muted, statut];
    return /* @__PURE__ */ <Badge color={color}>{label}</Badge>;
  };
  const badgeStatutFicheTexte = (statut) => (MAP_STATUTS_FICHE[statut] || [null, statut])[1];
  const badgeStatutSeance = (statut) => {
    const map = {
      EnAttente: [C.orange, "⏳ En attente"],
      Validee: [C.green, "✅ Validée"],
      Rejetee: [C.red, "✗ Rejetée"],
    };
    const [color, label] = map[statut] || [C.muted, statut];
    return /* @__PURE__ */ <Badge color={color}>{label}</Badge>;
  }; // Petit r\xE9capitulatif des 4 \xE9tapes du circuit, affich\xE9 en haut de la
  // fiche pour que chacun voie clairement o\xF9 elle en est et qui doit agir.
  const Circuit = ({ fiche }) => {
    const etapes = [
      {
        key: "Soumise",
        label: "Soumission (professeur)",
        fait: !!fiche.dateSoumission,
        qui: fiche.professeurNom,
        quand: fiche.dateSoumission,
      },
      {
        key: "ValideeIliass",
        label: "Validation des séances (Iliass Elmzioui)",
        fait: !!fiche.dateValidation,
        qui: fiche.valideePar,
        quand: fiche.dateValidation,
      },
      {
        key: "ReceptionAccusee",
        label: "Accusé de réception (Kawtar Bouabid)",
        fait: !!fiche.dateAccuseReception,
        qui: fiche.accuseReceptionPar,
        quand: fiche.dateAccuseReception,
      },
      {
        key: "Terminee",
        label: "Traitement du paiement (Fatima Zahra Zniber)",
        fait: !!fiche.dateTraitement,
        qui: fiche.traitePar,
        quand: fiche.dateTraitement,
      },
    ];
    return (
      /* @__PURE__ */ <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
        {etapes.map((e) => (
          /* @__PURE__ */ <div key={e.key} style={{ display: "flex", alignItems: "center", gap: 10, fontSize: 12.5 }}>
            <span
              style={{
                width: 20,
                height: 20,
                borderRadius: "50%",
                background: e.fait ? C.green : C.light,
                color: e.fait ? "white" : "#999",
                display: "flex",
                alignItems: "center",
                justifyContent: "center",
                fontSize: 11,
                fontWeight: 700,
                flexShrink: 0,
              }}
            >
              {e.fait ? "✓" : "–"}
            </span>
            <span style={{ fontWeight: e.fait ? 700 : 500, color: e.fait ? C.navy : C.muted }}>{e.label}</span>
            {e.fait && (
              /* @__PURE__ */ <span style={{ color: C.muted, fontSize: 11 }}>
                — {e.qui} · {formatDateFR(e.quand)}
              </span>
            )}
          </div>
        ))}
      </div>
    );
  };
  if (estProf) {
    const ficheOuverte = showFicheId ? fichesDeCeProf.find((f) => f.id === showFicheId) : null;
    if (ficheOuverte) {
      const seances = seancesDeFiche(ficheOuverte.id).sort((a, b) => (a.date || "").localeCompare(b.date || ""));
      const totalHeures = seances.reduce((s, x) => s + (x.nbHeures || 0), 0);
      const modulesOptions = (data.cours || []).filter((c) => c.professeurCin === monProfesseur?.cin);
      const toutesConfirmees = seances.length > 0 && seances.every((s) => s.confirmee);
      const periode = periodeDuMois(ficheOuverte.mois);
      return (
        /* @__PURE__ */ <div>
          <PageHeader
            title={`🕛 Pointage — ${formatMoisLabel(ficheOuverte.mois)}`}
            sub={`${seances.length} séance(s) · ${totalHeures}h au total · Période : du ${formatDateFR(periode.debut)} au ${formatDateFR(periode.fin)}`}
            action={
              /* @__PURE__ */ <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
                <Btn
                  variant="light"
                  onClick={() =>
                    genererFichePointage(ficheOuverte, seances, monProfesseur, formatMoisLabel, periodeDuMois)
                  }
                >
                  🖨️ Imprimer (PDF)
                </Btn>
                <Btn variant="light" onClick={handleRafraichirManuel}>
                  🔄 Actualiser
                </Btn>
                <Btn variant="light" onClick={() => setShowFicheId(null)}>
                  ← Retour
                </Btn>
              </div>
            }
          />
          {ficheOuverte.statut !== "Brouillon" && (
            /* @__PURE__ */ <Card style={{ marginBottom: 16, padding: 16 }}>
              <Circuit fiche={ficheOuverte} />
            </Card>
          )}
          {ficheOuverte.statut === "Brouillon" && (
            /* @__PURE__ */ <Card
              style={{
                marginBottom: 16,
                display: "flex",
                alignItems: "center",
                justifyContent: "space-between",
                flexWrap: "wrap",
                gap: 10,
              }}
            >
              <div>
                {badgeStatutFiche(ficheOuverte.statut)}
                {!toutesConfirmees && seances.length > 0 && (
                  /* @__PURE__ */ <span style={{ marginLeft: 10, fontSize: 12, color: C.orange, fontWeight: 600 }}>
                    Confirmez toutes les séances avant de soumettre
                  </span>
                )}
              </div>
              {seances.length > 0 && toutesConfirmees && (
                /* @__PURE__ */ <Btn onClick={() => handleSoumettreFiche(ficheOuverte)}>
                  → Soumettre la fiche pour validation
                </Btn>
              )}
            </Card>
          )}
          {ficheOuverte.statut === "Brouillon" && (
            /* @__PURE__ */ <Card style={{ marginBottom: 16, padding: 16 }}>
              <h4 style={{ margin: "0 0 12px", color: C.navy, fontSize: 14 }}>
                {editingSeanceId ? "Modifier la séance" : "Ajouter une séance"}
              </h4>
              <div
                style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(160px, 1fr))", gap: "0 14px" }}
              >
                <Input
                  label="Date *"
                  type="date"
                  value={formSeance.date}
                  min={periode?.debut}
                  max={periode?.fin}
                  onChange={(e) => setFormSeance((f) => ({ ...f, date: e.target.value }))}
                />
                <Sel
                  label="Module *"
                  value={formSeance.module}
                  onChange={(e) => setFormSeance((f) => ({ ...f, module: e.target.value }))}
                >
                  <option value="">-- Sélectionner --</option>
                  {modulesOptions.map((c) => (
                    /* @__PURE__ */ <option
                      key={c.id}
                      value={c.intitule}
                    >{`${c.intitule} (${filiereCode(c.filiere, data.filieres)} · ${c.semestre})`}</option>
                  ))}
                </Sel>
                <Input
                  label="Heure début *"
                  type="time"
                  value={formSeance.heureDebut}
                  onChange={(e) => setFormSeance((f) => ({ ...f, heureDebut: e.target.value }))}
                />
                <Input
                  label="Heure fin *"
                  type="time"
                  value={formSeance.heureFin}
                  onChange={(e) => setFormSeance((f) => ({ ...f, heureFin: e.target.value }))}
                />
              </div>
              <Input
                label="Commentaire (optionnel)"
                value={formSeance.commentaire}
                onChange={(e) => setFormSeance((f) => ({ ...f, commentaire: e.target.value }))}
                placeholder="Ex: Rattrapage suite à report, cours en demi-groupe..."
              />
              <div style={{ display: "flex", gap: 10 }}>
                <Btn onClick={() => handleEnregistrerSeance(ficheOuverte)}>
                  {editingSeanceId ? "Enregistrer les modifications" : "+ Ajouter la séance"}
                </Btn>
                {editingSeanceId && (
                  /* @__PURE__ */ <Btn variant="light" onClick={annulerEditionSeance}>
                    Annuler
                  </Btn>
                )}
              </div>
            </Card>
          )}
          <Card style={{ overflow: "hidden" }}>
            <DataTable
              cols={[
                { key: "date", label: "Date", render: (v) => formatDateFR(v) },
                { key: "module", label: "Module" },
                { key: "filiere", label: "Filière", render: (v) => filiereCode(v, data.filieres) },
                { key: "heureDebut", label: "Horaire", render: (_, row) => `${row.heureDebut} - ${row.heureFin}` },
                { key: "nbHeures", label: "Durée", render: (v) => `${v}h` },
                {
                  key: "commentaire",
                  label: "Commentaire",
                  render: (v) =>
                    v ? (
                      /* @__PURE__ */ <span style={{ fontSize: 11.5 }}>{v}</span>
                    ) : (
                      /* @__PURE__ */ <span style={{ color: C.muted, fontSize: 11 }}>—</span>
                    ),
                },
                ...(ficheOuverte.statut === "Brouillon"
                  ? [
                      {
                        key: "_confirmee",
                        label: "Confirmation",
                        render: (_, row) =>
                          row.confirmee ? (
                            /* @__PURE__ */ <Badge color={C.green}>✅ Confirmée</Badge>
                          ) : (
                            /* @__PURE__ */ <Btn small={true} variant="navy" onClick={() => handleConfirmerSeance(row)}>
                              ✓ Confirmer
                            </Btn>
                          ),
                      },
                    ]
                  : [{ key: "statut", label: "Statut", render: (v) => badgeStatutSeance(v) }]),
                ...(ficheOuverte.statut === "Brouillon"
                  ? [
                      {
                        key: "_actions",
                        label: "Actions",
                        render: (_, row) => (
                          /* @__PURE__ */ <div style={{ display: "flex", gap: 4 }}>
                            <IconBtn icon="✏️" title="Modifier" color={C.navy} onClick={() => openEditSeance(row)} />
                            <IconBtn
                              icon="🗑"
                              title="Supprimer"
                              color={C.red}
                              onClick={() => handleSupprimerSeance(row)}
                            />
                          </div>
                        ),
                      },
                    ]
                  : []),
              ]}
              rows={seances}
              emptyMsg="Aucune séance ajoutée"
            />
          </Card>
        </div>
      );
    }
    return (
      /* @__PURE__ */ <div>
        <PageHeader
          title="🕛 Mon pointage"
          sub={
            monProfesseur
              ? `${monProfesseur.prenom} ${monProfesseur.nom}`
              : "⚠️ Aucune fiche professeur associée à votre compte"
          }
          action={
            monProfesseur && (
              /* @__PURE__ */ <Btn onClick={() => setShowNouvelleFiche(true)}>+ Nouvelle fiche du mois</Btn>
            )
          }
        />
        {!monProfesseur && (
          /* @__PURE__ */ <Card style={{ padding: 24, textAlign: "center", color: C.muted, fontSize: 13 }}>
            <p style={{ margin: "0 0 10px" }}>
              Contactez l'administrateur pour associer votre compte à une fiche professeur.
            </p>
            <p style={{ margin: "0 0 4px", fontSize: 11.5 }}>
              Email connecté : <strong>{user?.email || "—"}</strong>
            </p>
            <p style={{ margin: "0 0 4px", fontSize: 11.5 }}>
              Nombre de fiches professeur visibles par votre compte : <strong>{(data.professeurs || []).length}</strong>
            </p>
            <p style={{ margin: 0, fontSize: 11.5 }}>
              Emails enregistrés dans la liste Professeurs : {(data.professeurs || [])
                .map((p) => p.email)
                .filter(Boolean)
                .join(", ") || "aucun"}
            </p>
          </Card>
        )}
        {monProfesseur &&
          (fichesDeCeProf.length === 0 ? (
            /* @__PURE__ */ <Card style={{ padding: 32, textAlign: "center", color: C.muted, fontSize: 13 }}>
              Aucune fiche de pointage pour le moment
            </Card>
          ) : (
            /* @__PURE__ */ <Card style={{ overflow: "hidden" }}>
              <DataTable
                cols={[
                  { key: "mois", label: "Mois", render: (v) => formatMoisLabel(v) },
                  { key: "statut", label: "Statut", render: (v) => badgeStatutFiche(v) },
                  { key: "id", label: "Séances", render: (_, row) => seancesDeFiche(row.id).length },
                  {
                    key: "id2",
                    label: "Heures",
                    render: (_, row) => `${seancesDeFiche(row.id).reduce((s, x) => s + (x.nbHeures || 0), 0)}h`,
                  },
                  {
                    key: "id3",
                    label: "",
                    render: (_, row) => (
                      /* @__PURE__ */ <Btn small={true} variant="navy" onClick={() => setShowFicheId(row.id)}>
                        Ouvrir
                      </Btn>
                    ),
                  },
                  {
                    key: "id4",
                    label: "Suppr.",
                    render: (_, row) =>
                      row.statut === "Brouillon" ? (
                        /* @__PURE__ */ <IconBtn
                          icon="🗑"
                          title="Supprimer"
                          color={C.red}
                          onClick={() => handleSupprimerFiche(row)}
                        />
                      ) : (
                        /* @__PURE__ */ <span style={{ color: C.muted, fontSize: 11 }}>—</span>
                      ),
                  },
                ]}
                rows={fichesDeCeProf}
                emptyMsg="Aucune fiche"
              />
            </Card>
          ))}
        {showNouvelleFiche && (
          /* @__PURE__ */ <Modal title="Nouvelle fiche de pointage" onClose={() => setShowNouvelleFiche(false)}>
            <Input
              label="Mois *"
              type="month"
              value={moisNouvelleFiche}
              onChange={(e) => setMoisNouvelleFiche(e.target.value)}
            />
            <div style={{ display: "flex", gap: 10, justifyContent: "flex-end", marginTop: 14 }}>
              <Btn variant="light" onClick={() => setShowNouvelleFiche(false)}>
                Annuler
              </Btn>
              <Btn onClick={handleCreerFiche}>Créer la fiche</Btn>
            </div>
          </Modal>
        )}
      </div>
    );
  }
  if (estValidateur) {
    const fichesFiltrees = (data.pointageFiches || [])
      .filter((f) => filtreValidateur === "Toutes" || f.statut === filtreValidateur)
      .sort((a, b) => b.mois.localeCompare(a.mois));
    const ficheOuverte = showFicheId ? (data.pointageFiches || []).find((f) => f.id === showFicheId) : null;
    if (ficheOuverte) {
      const seances = seancesDeFiche(ficheOuverte.id).sort((a, b) => (a.date || "").localeCompare(b.date || ""));
      const totalHeures = seances.reduce((s, x) => s + (x.nbHeures || 0), 0);
      const nbEnAttente = seances.filter((s) => s.statut === "EnAttente").length;
      return (
        /* @__PURE__ */ <div>
          <PageHeader
            title={`🕛 ${ficheOuverte.professeurNom} — ${formatMoisLabel(ficheOuverte.mois)}`}
            sub={`${seances.length} séance(s) · ${totalHeures}h au total`}
            action={
              /* @__PURE__ */ <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
                <Btn
                  variant="light"
                  onClick={() =>
                    genererFichePointage(
                      ficheOuverte,
                      seances,
                      (data.professeurs || []).find((p) => p.id === ficheOuverte.professeurId),
                      formatMoisLabel,
                      periodeDuMois
                    )
                  }
                >
                  🖨️ Imprimer (PDF)
                </Btn>
                <Btn variant="light" onClick={handleRafraichirManuel}>
                  🔄 Actualiser
                </Btn>
                <Btn variant="light" onClick={() => setShowFicheId(null)}>
                  ← Retour
                </Btn>
              </div>
            }
          />
          <Card style={{ marginBottom: 16, padding: 16 }}>
            <Circuit fiche={ficheOuverte} />
          </Card>
          <Card
            style={{
              marginBottom: 16,
              display: "flex",
              alignItems: "center",
              justifyContent: "space-between",
              flexWrap: "wrap",
              gap: 10,
            }}
          >
            {badgeStatutFiche(ficheOuverte.statut)}
            {ficheOuverte.statut === "Soumise" && nbEnAttente > 0 && estIliass && (
              /* @__PURE__ */ <span
                style={{ fontSize: 12, color: C.orange, fontWeight: 600 }}
              >{`${nbEnAttente} séance(s) en attente de validation`}</span>
            )}
            {ficheOuverte.statut === "Soumise" && seances.length === 0 && estIliass && (
              /* @__PURE__ */ <span style={{ fontSize: 12, color: C.red, fontWeight: 600 }}>
                ⚠️ Cette fiche ne contient aucune séance — rien à valider
              </span>
            )}
            {ficheOuverte.statut === "Soumise" && seances.length > 0 && nbEnAttente === 0 && estIliass && (
              /* @__PURE__ */ <Btn onClick={() => handleValiderFiche(ficheOuverte)}>
                → Marquer les séances validées — transmettre à Mme Bouabid
              </Btn>
            )}
            {ficheOuverte.statut === "Soumise" && !estIliass && (
              /* @__PURE__ */ <span style={{ fontSize: 12, color: C.muted }}>
                En attente de validation par M. Iliass Elmzioui
              </span>
            )}
            {ficheOuverte.statut === "ValideeIliass" && estKawtar && (
              /* @__PURE__ */ <Btn onClick={() => handleAccuserReception(ficheOuverte)}>
                📥 Accuser réception de la fiche
              </Btn>
            )}
            {ficheOuverte.statut === "ValideeIliass" && !estKawtar && (
              /* @__PURE__ */ <span style={{ fontSize: 12, color: C.muted }}>
                En attente d'accusé de réception par Mme Kawtar Bouabid
              </span>
            )}
            {ficheOuverte.statut === "ReceptionAccusee" && estFatimaZahra && (
              /* @__PURE__ */ <Btn onClick={() => handleConfirmerTraitement(ficheOuverte)}>
                💰 Confirmer le traitement du paiement et terminer
              </Btn>
            )}
            {ficheOuverte.statut === "ReceptionAccusee" && !estFatimaZahra && (
              /* @__PURE__ */ <span style={{ fontSize: 12, color: C.muted }}>
                En attente de traitement du paiement par Mme Fatima Zahra Zniber
              </span>
            )}
          </Card>
          <Card style={{ overflow: "hidden" }}>
            <DataTable
              cols={[
                { key: "date", label: "Date", render: (v) => formatDateFR(v) },
                { key: "module", label: "Module" },
                { key: "filiere", label: "Filière", render: (v) => filiereCode(v, data.filieres) },
                { key: "heureDebut", label: "Horaire", render: (_, row) => `${row.heureDebut} - ${row.heureFin}` },
                { key: "nbHeures", label: "Durée", render: (v) => `${v}h` },
                { key: "statut", label: "Statut", render: (v) => badgeStatutSeance(v) },
                {
                  key: "commentaire",
                  label: "Commentaire / Raison",
                  render: (v, row) =>
                    ficheOuverte.statut === "Soumise" && estIliass && row.statut !== "Validee" ? (
                      /* @__PURE__ */ <input
                        key={row.id}
                        defaultValue={v || ""}
                        placeholder="Ex: Motif du rejet, remarque..."
                        onBlur={(e) => {
                          if (e.target.value !== (v || "")) handleCommentaireAdminSeance(row, e.target.value);
                        }}
                        style={{
                          width: "100%",
                          minWidth: 160,
                          padding: "6px 8px",
                          border: `1px solid ${C.border}`,
                          borderRadius: 6,
                          fontSize: 11.5,
                          fontFamily: "inherit",
                        }}
                      />
                    ) : v ? (
                      /* @__PURE__ */ <span style={{ fontSize: 11.5 }}>{v}</span>
                    ) : (
                      /* @__PURE__ */ <span style={{ color: C.muted, fontSize: 11 }}>—</span>
                    ),
                },
                ...(ficheOuverte.statut === "Soumise" && estIliass
                  ? [
                      {
                        key: "_valider",
                        label: "Valider",
                        render: (_, row) =>
                          row.statut === "EnAttente" ? (
                            /* @__PURE__ */ <div style={{ display: "flex", gap: 4 }}>
                              <IconBtn
                                icon="✅"
                                title="Valider"
                                color={C.green}
                                onClick={() => handleValiderSeance(row, "Validee")}
                              />
                              <IconBtn
                                icon="✗"
                                title="Rejeter"
                                color={C.red}
                                onClick={() => handleValiderSeance(row, "Rejetee")}
                              />
                            </div>
                          ) : (
                            /* @__PURE__ */ <span style={{ fontSize: 11, color: C.muted }}>{row.valideePar}</span>
                          ),
                      },
                    ]
                  : []),
              ]}
              rows={seances}
              emptyMsg="Aucune séance"
            />
          </Card>
        </div>
      );
    }
    return (
      /* @__PURE__ */ <div>
        <PageHeader
          title="🕛 Pointage des enseignants vacataires"
          sub="Circuit : Soumission (professeur) → Validation (Iliass Elmzioui) → Accusé de réception (Kawtar Bouabid) → Traitement du paiement (Fatima Zahra Zniber)"
        />
        <div style={{ marginBottom: 16, display: "flex", gap: 8, flexWrap: "wrap" }}>
          {[
            ["Soumise", "⏳ À valider"],
            ["ValideeIliass", "📤 À réceptionner"],
            ["ReceptionAccusee", "💰 À traiter"],
            ["Terminee", "✅ Terminées"],
            ["Brouillon", "📝 Brouillons"],
            ["Toutes", "Toutes"],
          ].map(([s, label]) => (
            /* @__PURE__ */ <button
              key={s}
              onClick={() => setFiltreValidateur(s)}
              style={{
                padding: "8px 16px",
                borderRadius: 20,
                border: filtreValidateur === s ? "none" : `1px solid ${C.border}`,
                background: filtreValidateur === s ? C.navy : C.white,
                color: filtreValidateur === s ? "white" : "#555",
                cursor: "pointer",
                fontFamily: "inherit",
                fontWeight: 600,
                fontSize: 12.5,
              }}
            >
              {label}
            </button>
          ))}
        </div>
        {fichesFiltrees.length === 0 ? (
          /* @__PURE__ */ <Card style={{ padding: 32, textAlign: "center", color: C.muted, fontSize: 13 }}>
            Aucune fiche
          </Card>
        ) : (
          /* @__PURE__ */ <Card style={{ overflow: "hidden" }}>
            <DataTable
              cols={[
                { key: "professeurNom", label: "Professeur" },
                { key: "mois", label: "Mois", render: (v) => formatMoisLabel(v) },
                { key: "statut", label: "Statut", render: (v) => badgeStatutFiche(v) },
                { key: "id", label: "Séances", render: (_, row) => seancesDeFiche(row.id).length },
                {
                  key: "id2",
                  label: "Heures",
                  render: (_, row) => `${seancesDeFiche(row.id).reduce((s, x) => s + (x.nbHeures || 0), 0)}h`,
                },
                {
                  key: "id3",
                  label: "",
                  render: (_, row) => (
                    /* @__PURE__ */ <Btn small={true} variant="navy" onClick={() => setShowFicheId(row.id)}>
                      Ouvrir
                    </Btn>
                  ),
                },
                ...(role === "administrateur"
                  ? [
                      {
                        key: "id4",
                        label: "Suppr.",
                        render: (_, row) => (
                          /* @__PURE__ */ <IconBtn
                            icon="🗑"
                            title="Supprimer"
                            color={C.red}
                            onClick={() => handleSupprimerFiche(row)}
                          />
                        ),
                      },
                    ]
                  : []),
              ]}
              rows={fichesFiltrees}
              emptyMsg="Aucune fiche"
            />
          </Card>
        )}
      </div>
    );
  }
  return (
    /* @__PURE__ */ <Card style={{ padding: 32, textAlign: "center", color: C.muted, fontSize: 13 }}>
      Accès non autorisé
    </Card>
  );
}
function DemandesDocuments({ data, setData, role, user, toast }) {
  const canManage = ["administrateur", "directrice"].includes(role);
  const [showModal, setShowModal] = useState(false);
  const [search, setSearch] = useState("");
  const [form, setForm] = useState({ typeDocument: TYPES_DOCUMENT_DEMANDE[0], commentaire: "" });
  const monEtudiant = (data.etudiants || []).find(
    (e) =>
      (user?.numInscription && e.numInscription === user.numInscription) ||
      (user?.etudiantId && e.etudiantId === user.etudiantId)
  );
  const mesDemandes =
    role === "etudiant"
      ? (data.demandesDocuments || []).filter(
          (d) =>
            (user?.numInscription && d.numInscription === user.numInscription) ||
            (user?.etudiantId && d.etudiantId === user.etudiantId)
        )
      : data.demandesDocuments || [];
  const handleSubmit = () => {
    if (!user?.token || !monEtudiant) {
      toast("⚠️ Non connecté à SharePoint — demande non enregistrée");
      return;
    }
    const spF = {
      Title: `${monEtudiant.nom} ${monEtudiant.prenom} - ${form.typeDocument}`,
      EtudiantId: String(monEtudiant.id),
      NomEtudiant: `${monEtudiant.prenom} ${monEtudiant.nom}`,
      NumInscription: monEtudiant.numInscription || "",
      Filiere: monEtudiant.filiere || "",
      TypeDocument: form.typeDocument,
      DateDemande: today(),
      Statut: "En attente",
      Commentaire: form.commentaire || "",
    };
    saveToSP("ECOGEST_DemandesDocuments", spF, user.token).then((result) => {
      if (!result) {
        toast("⚠️ Échec de l'enregistrement dans SharePoint — demande non envoyée");
        return;
      }
      setData((d) => ({
        ...d,
        demandesDocuments: [
          ...d.demandesDocuments,
          {
            id: Date.now(),
            _spId: result.id,
            etudiantId: monEtudiant.id,
            etudiant: `${monEtudiant.prenom} ${monEtudiant.nom}`,
            numInscription: monEtudiant.numInscription,
            filiere: monEtudiant.filiere,
            typeDocument: form.typeDocument,
            dateDemande: today(),
            statut: "En attente",
            commentaire: form.commentaire || "",
            dateTraitement: "",
            traitePar: "",
          },
        ],
      }));
      toast("Demande envoyée dans SharePoint ✓");
    });
    setShowModal(false);
    setForm({ typeDocument: TYPES_DOCUMENT_DEMANDE[0], commentaire: "" });
  };
  const handleChangeStatut = (demande, nouveauStatut) => {
    if (!user?.token) {
      toast("⚠️ Non connecté à SharePoint — statut non modifié");
      return;
    }
    const updated = { ...demande, statut: nouveauStatut, dateTraitement: today(), traitePar: user?.nom || "" };
    setData((d) => ({ ...d, demandesDocuments: d.demandesDocuments.map((x) => (x.id === demande.id ? updated : x)) }));
    saveToSP(
      "ECOGEST_DemandesDocuments",
      { Statut: nouveauStatut, DateTraitement: today(), TraitePar: user?.nom || "" },
      user.token,
      demande._spId
    ).then((result) => {
      if (!result) toast("⚠️ Échec de la mise à jour du statut");
      else toast(`Statut mis à jour : ${nouveauStatut} ✓`);
    });
  };
  const handleDelete = (demande) => {
    if (!window.confirm(`Supprimer la demande de "${demande.etudiant}" (${demande.typeDocument}) ?`)) return;
    setData((d) => ({ ...d, demandesDocuments: d.demandesDocuments.filter((x) => x.id !== demande.id) }));
    if (user?.token && demande._spId) deleteFromSP("ECOGEST_DemandesDocuments", demande._spId, user.token);
    toast("Demande supprimée de SharePoint ✓");
  };
  const ordreStatuts = STATUTS_DEMANDE;
  const demandesFiltrees = canManage
    ? mesDemandes.filter(
        (d) =>
          !search ||
          `${d.etudiant || ""} ${d.numInscription || ""} ${d.typeDocument || ""}`
            .toLowerCase()
            .includes(search.toLowerCase())
      )
    : mesDemandes;
  const parStatut = {};
  demandesFiltrees.forEach((d) => {
    const s = d.statut || "En attente";
    if (!parStatut[s]) parStatut[s] = [];
    parStatut[s].push(d);
  });
  Object.keys(parStatut).forEach((s) =>
    parStatut[s].sort((a, b) => (b.dateDemande || "").localeCompare(a.dateDemande || ""))
  );
  return (
    /* @__PURE__ */ <div>
      <PageHeader
        title={role === "etudiant" ? "Mes demandes de documents" : "Demandes de documents administratifs"}
        sub={
          role === "etudiant"
            ? `${mesDemandes.length} demande(s)`
            : `Liste SharePoint "${SP.lists.demandesDocuments}" — ${data.demandesDocuments.length} demande(s)`
        }
        action={role === "etudiant" && /* @__PURE__ */ <Btn onClick={() => setShowModal(true)}>+ Nouvelle demande</Btn>}
      />
      {canManage && (
        /* @__PURE__ */ <div style={{ display: "flex", gap: 10, marginBottom: 14 }}>
          <input
            placeholder="🔍  Nom, N° inscription, type de document..."
            value={search}
            onChange={(e) => setSearch(e.target.value)}
            style={{
              flex: 1,
              padding: "10px 14px",
              border: `1px solid ${C.border}`,
              borderRadius: 8,
              fontSize: 13,
              boxSizing: "border-box",
              fontFamily: "inherit",
            }}
          />
        </div>
      )}
      {ordreStatuts.filter((s) => parStatut[s]?.length).length === 0 ? (
        /* @__PURE__ */ <Card style={{ padding: 32, textAlign: "center", color: C.muted, fontSize: 13 }}>
          Aucune demande trouvée
        </Card>
      ) : (
        ordreStatuts
          .filter((s) => parStatut[s]?.length)
          .map((statut) => (
            /* @__PURE__ */ <div key={statut} style={{ marginBottom: 20 }}>
              <div
                style={{
                  fontSize: 13,
                  fontWeight: 800,
                  textTransform: "uppercase",
                  letterSpacing: "0.05em",
                  marginBottom: 10,
                  padding: "5px 12px",
                  borderRadius: 20,
                  display: "inline-block",
                  background: STATUT_COLOR[statut] + "22",
                  color: STATUT_COLOR[statut],
                  border: `1px solid ${STATUT_COLOR[statut]}55`,
                }}
              >
                {statut} <span style={{ fontWeight: 700 }}>({parStatut[statut].length})</span>
              </div>
              <Card style={{ overflow: "hidden" }}>
                <DataTable
                  cols={[
                    { key: "etudiant", label: "Étudiant" },
                    { key: "filiere", label: "Filière" },
                    { key: "typeDocument", label: "Document demandé" },
                    { key: "dateDemande", label: "Date de la demande", render: (v) => formatDateFR(v) },
                    { key: "commentaire", label: "Commentaire" },
                    ...(canManage
                      ? [
                          {
                            key: "_statut",
                            label: "Changer le statut",
                            render: (_, row) => (
                              /* @__PURE__ */ <select
                                value={row.statut}
                                onChange={(e) => handleChangeStatut(row, e.target.value)}
                                style={{
                                  padding: "5px 8px",
                                  border: `1px solid ${C.border}`,
                                  borderRadius: 6,
                                  fontSize: 12,
                                  fontFamily: "inherit",
                                }}
                              >
                                {STATUTS_DEMANDE.map((s) => (
                                  /* @__PURE__ */ <option key={s} value={s}>
                                    {s}
                                  </option>
                                ))}
                              </select>
                            ),
                          },
                        ]
                      : []),
                    ...(canManage
                      ? [
                          {
                            key: "_suppr",
                            label: "Suppr.",
                            render: (_, row) => (
                              /* @__PURE__ */ <Btn small={true} variant="danger" onClick={() => handleDelete(row)}>
                                🗑
                              </Btn>
                            ),
                          },
                        ]
                      : []),
                  ]}
                  rows={parStatut[statut]}
                  emptyMsg="Aucune demande trouvée"
                />
              </Card>
            </div>
          ))
      )}
      {showModal && (
        /* @__PURE__ */ <Modal title="Nouvelle demande de document → SharePoint" onClose={() => setShowModal(false)}>
          <Sel
            label="Document souhaité *"
            value={form.typeDocument}
            onChange={(e) => setForm((f) => ({ ...f, typeDocument: e.target.value }))}
          >
            {TYPES_DOCUMENT_DEMANDE.map((t) => (
              /* @__PURE__ */ <option key={t}>{t}</option>
            ))}
          </Sel>
          <div style={{ marginBottom: 13 }}>
            <label style={{ fontSize: 12, fontWeight: 600, color: "#555", display: "block", marginBottom: 4 }}>
              Commentaire (optionnel)
            </label>
            <textarea
              value={form.commentaire}
              onChange={(e) => setForm((f) => ({ ...f, commentaire: e.target.value }))}
              rows={3}
              placeholder="Précisez votre demande si besoin..."
              style={{
                width: "100%",
                padding: "9px 12px",
                border: `1px solid ${C.border}`,
                borderRadius: 8,
                fontSize: 13,
                fontFamily: "inherit",
                boxSizing: "border-box",
                resize: "vertical",
              }}
            />
          </div>
          <div style={{ display: "flex", gap: 10, justifyContent: "flex-end" }}>
            <Btn variant="light" onClick={() => setShowModal(false)}>
              Annuler
            </Btn>
            <Btn onClick={handleSubmit}>Envoyer la demande → SharePoint</Btn>
          </div>
        </Modal>
      )}
    </div>
  );
}
const JOURS_SEMAINE = ["Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi", "Samedi", "Dimanche"];
const JOURS_GROUPE_JOUR = ["Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi"];
const JOURS_GROUPE_SOIR = ["Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi", "Samedi", "Dimanche"];
const SEANCES_RAMADAN_JOUR = ["12h00 - 13h30", "13h40 - 15h10"];
const SEANCES_RAMADAN_SOIR = ["11h00 - 12h30", "13h00 - 14h30"];
const JOURS_RAMADAN_SOIR = ["Samedi", "Dimanche"];
const MS_JOUR = 864e5;
const toISODate = (d) => d.toISOString().slice(0, 10);
const parseISODate = (s) => /* @__PURE__ */ new Date(s + "T00:00:00");
const addJoursISO = (iso, n) => toISODate(new Date(parseISODate(iso).getTime() + n * MS_JOUR));
const lundiDeLaSemaineISO = (iso) => {
  const d = parseISODate(iso);
  const jour = d.getDay();
  const decalage = jour === 0 ? -6 : 1 - jour;
  return addJoursISO(iso, decalage);
};
const genererSemaines = (debutISO, finISO) => {
  if (!debutISO || !finISO) return [];
  const semaines = [];
  let cur = lundiDeLaSemaineISO(debutISO);
  const finLundi = lundiDeLaSemaineISO(finISO);
  let garde = 0;
  while (cur <= finLundi && garde < 60) {
    semaines.push(cur);
    cur = addJoursISO(cur, 7);
    garde++;
  }
  return semaines;
};
function EmploiDuTemps({ data, setData, role, user, toast }) {
  const JOUR_COLORS = {
    Lundi: "#2563eb",
    Mardi: "#0e9187",
    Mercredi: "#7c3aed",
    Jeudi: "#e56b2d",
    Vendredi: "#c2185b",
    Samedi: "#059669",
    Dimanche: "#64748b",
  };
  const GRADIENTS_MODULES = [
    ["#0b1f3f", "#2f5fa8"],
    ["#0a3d3a", "#0e9187"],
    ["#2c1653", "#7c3aed"],
    ["#5c2405", "#e56b2d"],
    ["#4a0d24", "#c2185b"],
    ["#0c2e1c", "#059669"],
    ["#450a0a", "#c0392b"],
    ["#1e2533", "#546178"],
    ["#3d2a06", "#c99a2e"],
    ["#0b2942", "#0891b2"],
  ];
  const colorForModule = (nom) => {
    const s = String(nom || "");
    let hash = 0;
    for (let i = 0; i < s.length; i++) hash = (hash * 31 + s.charCodeAt(i)) >>> 0;
    return GRADIENTS_MODULES[hash % GRADIENTS_MODULES.length];
  };
  const canEdit = ["administrateur", "directrice"].includes(role);
  const [showModal, setShowModal] = useState(false);
  const [editItem, setEditItem] = useState(null);
  const [showDeleteChoice, setShowDeleteChoice] = useState(null);
  const emptyForm = {
    filiere: "",
    semestre: "",
    jour: "",
    seance: "",
    module: "",
    typeRecurrence: "Hebdomadaire",
    dateDebut: "",
    dateFin: "",
    estRattrapage: false,
    lienTeams: "",
    lienEnregistrements: "",
  };
  const [form, setForm] = useState(emptyForm);
  const monEtudiant = (data.etudiants || []).find(
    (e) =>
      (user?.numInscription && e.numInscription === user.numInscription) ||
      (user?.etudiantId && e.etudiantId === user.etudiantId)
  );
  const filieresOptions = getFilieres(data);
  const [filterFiliere, setFilterFiliere] = useState(
    role === "etudiant" ? monEtudiant?.filiere || "" : filieresOptions[0] || ""
  );
  const semestresDisponibles = [
    .../* @__PURE__ */ new Set([
      ...SEMESTRES_MODULE_SEED,
      ...(data.cours || []).map((c) => c.semestre).filter(Boolean),
      ...(data.emploiDuTemps || []).map((e) => e.semestre).filter(Boolean),
    ]),
  ].sort((a, b) => a.localeCompare(b, "fr", { numeric: true, sensitivity: "base" }));
  const [filterSemestre, setFilterSemestre] = useState(semestresDisponibles[0] || "");
  const semestreInfo = (data.semestresDates || []).find((s) => s.semestre === filterSemestre);
  const filiereInfo = (data.filieres || []).find((f) => f.intitule === filterFiliere);
  const typeSeancesGroupe = filiereInfo?.typeSeances || "Soir";
  const semainesDisponibles = genererSemaines(semestreInfo?.dateDebut, semestreInfo?.dateFin);
  const calculerSemaineParDefaut = () => {
    const auj = lundiDeLaSemaineISO(toISODate(/* @__PURE__ */ new Date()));
    if (semainesDisponibles.length === 0) return auj;
    if (semainesDisponibles.includes(auj)) return auj;
    return auj < semainesDisponibles[0] ? semainesDisponibles[0] : semainesDisponibles[semainesDisponibles.length - 1];
  };
  const [semaineSelectionnee, setSemaineSelectionnee] = useState(calculerSemaineParDefaut());
  useEffect(() => {
    if (semainesDisponibles.length && !semainesDisponibles.includes(semaineSelectionnee)) {
      setSemaineSelectionnee(calculerSemaineParDefaut());
    }
  }, [filterSemestre]);
  const { ramadanDebut, ramadanFin } = data.configuration || {};
  const estSemaineRamadan = !!(
    ramadanDebut &&
    ramadanFin &&
    addJoursISO(semaineSelectionnee, 6) >= ramadanDebut &&
    semaineSelectionnee <= ramadanFin
  );
  const seancesAffichees = estSemaineRamadan
    ? typeSeancesGroupe === "Jour"
      ? SEANCES_RAMADAN_JOUR
      : SEANCES_RAMADAN_SOIR
    : typeSeancesGroupe === "Jour"
      ? SEANCES_HORAIRE.filter((s) => s !== "19h00 - 21h00")
      : SEANCES_HORAIRE.filter((s) => s === "19h00 - 21h00");
  const joursAffiches = estSemaineRamadan
    ? typeSeancesGroupe === "Jour"
      ? JOURS_GROUPE_JOUR
      : JOURS_RAMADAN_SOIR
    : typeSeancesGroupe === "Jour"
      ? JOURS_GROUPE_JOUR
      : JOURS_GROUPE_SOIR;
  const dateDuJour = (jour) => addJoursISO(semaineSelectionnee, JOURS_SEMAINE.indexOf(jour));
  const modulesOptions = (data.cours || []).filter(
    (c) =>
      (!form.filiere || (c.filieresListe || [c.filiere]).includes(form.filiere)) &&
      (!form.semestre || c.semestre === form.semestre)
  );
  const moduleChoisi = (data.cours || []).find((c) => c.intitule === form.module);
  const openAdd = (jour, seance) => {
    setEditItem(null);
    const d = jour ? dateDuJour(jour) : "";
    setForm({
      ...emptyForm,
      filiere: filterFiliere,
      semestre: filterSemestre,
      jour: jour || "",
      seance: seance || "",
      dateDebut: d,
      dateFin: semestreInfo?.dateFin || "",
    });
    setShowModal(true);
  };
  const openEdit = (item) => {
    setEditItem(item);
    setForm({
      filiere: item.filiere,
      semestre: item.semestre,
      jour: item.jour,
      seance: item.seance,
      module: item.module,
      typeRecurrence: item.typeRecurrence || "Hebdomadaire",
      dateDebut: item.dateDebut || "",
      dateFin: item.dateFin || "",
      estRattrapage: !!item.estRattrapage,
      lienTeams: item.lienTeams || "",
      lienEnregistrements: item.lienEnregistrements || "",
    });
    setShowModal(true);
  };
  const handleSave = () => {
    if (!user?.token) {
      toast("⚠️ Non connecté à SharePoint — créneau non enregistré");
      setShowModal(false);
      return;
    }
    const typeRecurrenceEffectif = form.estRattrapage ? "Unique" : form.typeRecurrence || "Hebdomadaire";
    const spF = {
      Title: `${form.filiere} - ${form.jour} - ${form.seance}`,
      Filiere: form.filiere || "",
      Semestre: form.semestre || "",
      JourSemaine: form.jour || "",
      Seance: form.seance || "",
      Module: form.module || "",
      Professeur: moduleChoisi?.professeur || "",
      ProfesseurCIN: moduleChoisi?.professeurCin || "",
      Salle: moduleChoisi?.salle || "",
      TypeRecurrence: typeRecurrenceEffectif,
      DateDebut: form.dateDebut || "",
      DateFin: typeRecurrenceEffectif === "Unique" ? form.dateDebut || "" : form.dateFin || "",
      DatesExclues: editItem ? (editItem.datesExclues || []).join(", ") : "",
      EstRattrapage: form.estRattrapage ? "Oui" : "Non",
      LienTeams: form.lienTeams || "",
      LienEnregistrements: form.lienEnregistrements || "",
    };
    setFilterFiliere(form.filiere);
    setFilterSemestre(form.semestre);
    if (editItem) {
      const updated = {
        ...form,
        id: editItem.id,
        _spId: editItem._spId,
        professeur: moduleChoisi?.professeur || "",
        professeurCin: moduleChoisi?.professeurCin || "",
        typeRecurrence: typeRecurrenceEffectif,
        dateFin: typeRecurrenceEffectif === "Unique" ? form.dateDebut : form.dateFin,
        datesExclues: editItem.datesExclues || [],
      };
      setData((d) => ({ ...d, emploiDuTemps: d.emploiDuTemps.map((x) => (x.id === editItem.id ? updated : x)) }));
      saveToSP("ECOGEST_EmploiDuTemps", spF, user.token, editItem._spId).then((result) => {
        if (!result) {
          toast("⚠️ Échec de la mise à jour dans SharePoint");
          return;
        }
        toast("Créneau mis à jour dans SharePoint ✓");
      });
    } else {
      const newItem = {
        ...form,
        id: Date.now(),
        professeur: moduleChoisi?.professeur || "",
        professeurCin: moduleChoisi?.professeurCin || "",
        typeRecurrence: typeRecurrenceEffectif,
        dateFin: typeRecurrenceEffectif === "Unique" ? form.dateDebut : form.dateFin,
        datesExclues: [],
      };
      saveToSP("ECOGEST_EmploiDuTemps", spF, user.token).then(async (result) => {
        if (!result) {
          toast("⚠️ Échec de l'enregistrement dans SharePoint — créneau non ajouté");
          return;
        }
        setData((d) => ({ ...d, emploiDuTemps: [...d.emploiDuTemps, { ...newItem, _spId: result.id }] }));
        const autresFilieres = form.estRattrapage
          ? []
          : (moduleChoisi?.filieresListe || []).filter((f) => f && f !== form.filiere);
        if (autresFilieres.length > 0) {
          const propages = [];
          for (const autreFiliere of autresFilieres) {
            const dejaExiste = (data.emploiDuTemps || []).some(
              (e) =>
                e.filiere === autreFiliere &&
                e.semestre === form.semestre &&
                e.jour === form.jour &&
                e.seance === form.seance &&
                e.module === form.module
            );
            if (dejaExiste) continue;
            const spF2 = { ...spF, Title: `${autreFiliere} - ${form.jour} - ${form.seance}`, Filiere: autreFiliere };
            const r2 = await saveToSP("ECOGEST_EmploiDuTemps", spF2, user.token);
            if (r2) propages.push({ ...newItem, id: Date.now() + Math.random(), _spId: r2.id, filiere: autreFiliere });
          }
          if (propages.length > 0) {
            setData((d) => ({ ...d, emploiDuTemps: [...d.emploiDuTemps, ...propages] }));
            toast(`Créneau enregistré ✓ — module commun, propagé à ${propages.length} autre(s) filière(s)`);
            return;
          }
        }
        toast("Créneau enregistré dans SharePoint ✓");
      });
    }
    setShowModal(false);
    setForm(emptyForm);
  };
  const requestDelete = (item, jour) => {
    if (item.typeRecurrence === "Unique") {
      if (!window.confirm(`Supprimer cette séance unique de "${item.module}" ?`)) return;
      deleteSerieEntiere(item);
      return;
    }
    setShowDeleteChoice({ item, jour });
  };
  const deleteSerieEntiere = (item) => {
    setData((d) => ({ ...d, emploiDuTemps: d.emploiDuTemps.filter((x) => x.id !== item.id) }));
    if (user?.token && item._spId) deleteFromSP("ECOGEST_EmploiDuTemps", item._spId, user.token);
    toast("Créneau supprimé de SharePoint ✓");
  };
  const deleteOccurrenceSeule = (item, jour) => {
    const dateExclue = dateDuJour(jour);
    const nouvellesExclusions = [.../* @__PURE__ */ new Set([...(item.datesExclues || []), dateExclue])];
    const updated = { ...item, datesExclues: nouvellesExclusions };
    setData((d) => ({ ...d, emploiDuTemps: d.emploiDuTemps.map((x) => (x.id === item.id ? updated : x)) }));
    if (user?.token && item._spId)
      saveToSP("ECOGEST_EmploiDuTemps", { DatesExclues: nouvellesExclusions.join(", ") }, user.token, item._spId);
    toast(`Occurrence du ${formatDateFR(dateExclue)} annulée ✓ (série conservée)`);
  };
  const grille = (data.emploiDuTemps || []).filter((e) => e.filiere === filterFiliere && e.semestre === filterSemestre);
  const appliqueASemaine = (slot, jour) => {
    const d = dateDuJour(jour);
    if (slot.typeRecurrence === "Unique") return slot.dateDebut === d;
    if ((slot.datesExclues || []).includes(d)) return false;
    if (slot.dateDebut && d < slot.dateDebut) return false;
    if (slot.dateFin && d > slot.dateFin) return false;
    return true;
  };
  const cellsFor = (jour, seance) =>
    grille.filter((e) => e.jour === jour && e.seance === seance && appliqueASemaine(e, jour));
  const modulesDuGroupe = (data.cours || []).filter(
    (c) => (c.filieresListe || [c.filiere]).includes(filterFiliere) && c.semestre === filterSemestre
  );
  const handleSetDate = (module, champ, valeur) => {
    if (!user?.token) {
      toast("⚠️ Non connecté à SharePoint — date non enregistrée");
      return;
    }
    const updated = { ...module, [champ]: valeur };
    setData((d) => ({ ...d, cours: d.cours.map((c) => (c.id === module.id ? updated : c)) }));
    saveToSP(
      "ECOGEST_Cours",
      { [champ === "dateCC" ? "DateCC" : "DateExamen"]: valeur },
      user.token,
      module._spId
    ).then((result) => {
      if (!result) {
        toast("⚠️ Échec de l'enregistrement de la date");
        return;
      }
      toast("Date enregistrée dans SharePoint ✓");
    });
    const label = champ === "dateCC" ? "Contrôle continu" : "Examen";
    const titre = `${label} - ${module.intitule}`;
    const existant = (data.calendrier || []).find((e) => e.titre === titre && e.filiere === module.filiere);
    const spCal = {
      Title: titre,
      TypeEvenement: "Examen",
      DateDebut: valeur,
      Description: `${label} du module ${module.intitule} (${module.semestre || ""})`,
      Filiere: module.filiere || "",
    };
    if (!valeur) return;
    if (existant) {
      setData((d) => ({
        ...d,
        calendrier: d.calendrier.map((e) => (e.id === existant.id ? { ...e, dateDebut: valeur } : e)),
      }));
      saveToSP("ECOGEST_Calendrier", spCal, user.token, existant._spId);
    } else {
      saveToSP("ECOGEST_Calendrier", spCal, user.token).then((result) => {
        if (result)
          setData((d) => ({
            ...d,
            calendrier: [
              ...d.calendrier,
              {
                id: Date.now(),
                _spId: result.id,
                titre,
                type: "Examen",
                dateDebut: valeur,
                dateFin: "",
                description: spCal.Description,
                filiere: module.filiere,
              },
            ],
          }));
      });
    }
  };
  const joursRestants = (dateStr) => {
    if (!dateStr) return null;
    const diff = Math.ceil((new Date(dateStr) - new Date(/* @__PURE__ */ new Date().toDateString())) / 864e5);
    if (diff < 0) return { texte: "Passé", color: C.muted };
    if (diff === 0) return { texte: "Aujourd'hui", color: C.red };
    if (diff <= 7) return { texte: `Dans ${diff}j`, color: C.red };
    if (diff <= 14) return { texte: `Dans ${diff}j`, color: C.orange };
    return { texte: `Dans ${diff}j`, color: C.green };
  };
  return (
    /* @__PURE__ */ <div>
      <PageHeader
        title={role === "etudiant" ? "Mon emploi du temps" : "Emploi du temps"}
        sub={
          role === "etudiant"
            ? monEtudiant?.filiere || ""
            : `Liste SharePoint "${SP.lists.emploiDuTemps}" — ${(data.emploiDuTemps || []).length} créneaux`
        }
        action={canEdit && /* @__PURE__ */ <Btn onClick={() => openAdd()}>+ Ajouter un créneau</Btn>}
      />
      <div style={{ marginBottom: 16, display: "flex", gap: 10, flexWrap: "wrap" }}>
        {role !== "etudiant" && (
          /* @__PURE__ */ <select
            value={filterFiliere}
            onChange={(e) => setFilterFiliere(e.target.value)}
            style={{
              padding: "9px 14px",
              border: `1px solid ${C.border}`,
              borderRadius: 8,
              fontSize: 13,
              fontFamily: "inherit",
              background: C.white,
              minWidth: 260,
            }}
          >
            <option value="">-- Sélectionner un groupe --</option>
            {filieresOptions.map((f) => (
              /* @__PURE__ */ <option key={f} value={f}>
                {f}
              </option>
            ))}
          </select>
        )}
        <div style={{ display: "flex", gap: 6 }}>
          {semestresDisponibles.map((s) => (
            /* @__PURE__ */ <button
              key={s}
              onClick={() => setFilterSemestre(s)}
              style={{
                padding: "9px 16px",
                borderRadius: 8,
                border: `1px solid ${filterSemestre === s ? C.navy : C.border}`,
                background: filterSemestre === s ? C.navy : C.white,
                color: filterSemestre === s ? C.white : "#555",
                fontSize: 13,
                fontWeight: 600,
                cursor: "pointer",
                fontFamily: "inherit",
              }}
            >
              {s}
            </button>
          ))}
        </div>
      </div>
      {filterFiliere && (
        /* @__PURE__ */ <div style={{ marginBottom: 14, display: "flex", gap: 8, flexWrap: "wrap" }}>
          <span
            style={{
              display: "inline-flex",
              alignItems: "center",
              gap: 6,
              fontSize: 12,
              fontWeight: 700,
              padding: "5px 12px",
              borderRadius: 20,
              background: typeSeancesGroupe === "Jour" ? "#fff6e8" : C.navyLight,
              color: typeSeancesGroupe === "Jour" ? "#a5690a" : C.navy,
              border: `1px solid ${typeSeancesGroupe === "Jour" ? "#f5c98a" : C.navy + "30"}`,
            }}
          >
            {typeSeancesGroupe === "Jour" ? "☀️ Groupe de journée" : "🌙 Groupe du soir"}
            {estSemaineRamadan
              ? typeSeancesGroupe === "Jour"
                ? " — Lundi-Vendredi, 12h00-13h30 / 13h40-15h10"
                : " — Samedi-Dimanche, 11h00-12h30 / 13h00-14h30"
              : typeSeancesGroupe === "Jour"
                ? " — tous les créneaux sauf 19h00-21h00"
                : " — créneau 19h00-21h00 uniquement"}
          </span>
          {estSemaineRamadan && (
            /* @__PURE__ */ <span
              style={{
                display: "inline-flex",
                alignItems: "center",
                gap: 6,
                fontSize: 12,
                fontWeight: 700,
                padding: "5px 12px",
                borderRadius: 20,
                background: "#f6f0fb",
                color: "#7c3aed",
                border: "1px solid #d9c3ec",
              }}
            >
              🌙 Semaine de Ramadan — horaires adaptés
            </span>
          )}
        </div>
      )}
      {filterFiliere &&
        (semainesDisponibles.length === 0 ? (
          /* @__PURE__ */ <div
            style={{
              marginBottom: 16,
              background: "#fff6e8",
              border: "1px solid #f5c98a",
              borderRadius: 8,
              padding: "10px 14px",
              fontSize: 12.5,
              color: "#a5690a",
            }}
          >
            ⚠️ Les dates de début/fin de "{filterSemestre}" ne sont pas encore renseignées dans Paramétrage — impossible
            de générer les semaines. La grille ci-dessous affiche la semaine en cours par défaut.
          </div>
        ) : (
          /* @__PURE__ */ <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 16 }}>
            <button
              onClick={() => {
                const idx = semainesDisponibles.indexOf(semaineSelectionnee);
                if (idx > 0) setSemaineSelectionnee(semainesDisponibles[idx - 1]);
              }}
              disabled={semainesDisponibles.indexOf(semaineSelectionnee) <= 0}
              style={{
                border: `1px solid ${C.border}`,
                background: C.white,
                borderRadius: 8,
                padding: "8px 12px",
                cursor: "pointer",
                fontSize: 14,
                opacity: semainesDisponibles.indexOf(semaineSelectionnee) <= 0 ? 0.4 : 1,
              }}
            >
              ◀
            </button>
            <select
              value={semaineSelectionnee}
              onChange={(e) => setSemaineSelectionnee(e.target.value)}
              style={{
                padding: "9px 14px",
                border: `1px solid ${C.navy}`,
                borderRadius: 8,
                fontSize: 13,
                fontWeight: 700,
                color: C.navy,
                fontFamily: "inherit",
                background: C.navyLight,
                minWidth: 260,
              }}
            >
              {semainesDisponibles.map((sem) => (
                /* @__PURE__ */ <option key={sem} value={sem}>
                  Semaine du {formatDateFR(sem)} au {formatDateFR(addJoursISO(sem, 5))}
                </option>
              ))}
            </select>
            <button
              onClick={() => {
                const idx = semainesDisponibles.indexOf(semaineSelectionnee);
                if (idx < semainesDisponibles.length - 1) setSemaineSelectionnee(semainesDisponibles[idx + 1]);
              }}
              disabled={semainesDisponibles.indexOf(semaineSelectionnee) >= semainesDisponibles.length - 1}
              style={{
                border: `1px solid ${C.border}`,
                background: C.white,
                borderRadius: 8,
                padding: "8px 12px",
                cursor: "pointer",
                fontSize: 14,
                opacity: semainesDisponibles.indexOf(semaineSelectionnee) >= semainesDisponibles.length - 1 ? 0.4 : 1,
              }}
            >
              ▶
            </button>
          </div>
        ))}
      {!filterFiliere ? (
        /* @__PURE__ */ <Card style={{ padding: 32, textAlign: "center", color: C.muted, fontSize: 13 }}>
          Sélectionnez un groupe pour afficher son emploi du temps
        </Card>
      ) : (
        /* @__PURE__ */ <div
          style={{
            background: "linear-gradient(160deg, #d9e8f9 0%, #eef4fb 40%, #ffffff 100%)",
            borderRadius: 18,
            padding: 3,
            boxShadow: "0 10px 30px rgba(20,40,80,0.12)",
          }}
        >
          <Card
            style={{
              overflow: "auto",
              padding: 0,
              borderRadius: 16,
              border: "none",
              boxShadow: "none",
              background: "transparent",
            }}
          >
            <table style={{ width: "100%", borderCollapse: "collapse", minWidth: 720 }}>
              <thead>
                <tr
                  style={{
                    background: `linear-gradient(120deg, ${C.navy} 0%, #1d5f8a 45%, ${C.teal} 100%)`,
                    boxShadow: "0 4px 10px rgba(10,25,55,0.25)",
                  }}
                >
                  <th
                    style={{
                      padding: "14px 16px",
                      color: "white",
                      fontSize: 11.5,
                      fontWeight: 700,
                      textAlign: "left",
                      whiteSpace: "nowrap",
                    }}
                  >
                    🕓 Séance
                  </th>
                  {joursAffiches.map((j, ji) => (
                    /* @__PURE__ */ <th
                      key={j}
                      style={{
                        padding: "14px 10px",
                        color: "white",
                        fontSize: 11.5,
                        fontWeight: 700,
                        textAlign: "center",
                        borderLeft: ji > 0 ? "1px solid rgba(255,255,255,0.18)" : "none",
                      }}
                    >
                      {j}
                      <div style={{ fontSize: 9.5, fontWeight: 500, opacity: 0.88, marginTop: 1 }}>
                        {formatDateFR(dateDuJour(j))}
                      </div>
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {seancesAffichees.map((seance, ri) => (
                  /* @__PURE__ */ <tr key={seance} style={{ borderBottom: `1px solid ${C.border}` }}>
                    <td
                      style={{
                        padding: "12px 14px",
                        fontSize: 11.5,
                        fontWeight: 800,
                        color: C.navy,
                        whiteSpace: "nowrap",
                        borderRight: `2px solid ${C.navy}18`,
                        background: ri % 2 === 0 ? "#ffffff" : "#f5f8fc",
                      }}
                    >
                      {seance}
                    </td>
                    {joursAffiches.map((jour) => {
                      const cellItems = cellsFor(jour, seance);
                      return (
                        /* @__PURE__ */ <td
                          key={jour}
                          style={{
                            padding: 6,
                            textAlign: "center",
                            verticalAlign: "middle",
                            minWidth: 130,
                            background:
                              ri % 2 === 0
                                ? "linear-gradient(160deg, #eef4fb 0%, #f7f9fc 100%)"
                                : "linear-gradient(160deg, #f5f8fc 0%, #ffffff 100%)",
                            borderLeft: `1px solid ${C.border}`,
                          }}
                        >
                          <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
                            {cellItems.map((cell) => (
                              /* @__PURE__ */ <div
                                key={cell.id}
                                onClick={() => canEdit && openEdit(cell)}
                                onMouseEnter={(e) => {
                                  e.currentTarget.style.boxShadow = "0 6px 16px rgba(20,30,60,0.16)";
                                  e.currentTarget.style.transform = "translateY(-1px)";
                                }}
                                onMouseLeave={(e) => {
                                  e.currentTarget.style.boxShadow = "0 1px 3px rgba(20,30,60,0.06)";
                                  e.currentTarget.style.transform = "translateY(0)";
                                }}
                                style={{
                                  background: `linear-gradient(135deg, ${colorForModule(cell.id)[0]} 0%, ${colorForModule(cell.id)[1]} 100%)`,
                                  border: "1px solid rgba(255,255,255,0.12)",
                                  borderLeft: `5px solid rgba(255,255,255,0.55)`,
                                  borderRadius: 10,
                                  padding: "9px 10px",
                                  cursor: canEdit ? "pointer" : "default",
                                  position: "relative",
                                  textAlign: "left",
                                  boxShadow: "0 3px 10px rgba(10,15,35,0.28)",
                                  transition: "box-shadow 0.15s, transform 0.15s",
                                }}
                              >
                                <div style={{ fontSize: 11.5, fontWeight: 800, color: "#ffffff" }}>{cell.module}</div>
                                {cell.professeur && (
                                  /* @__PURE__ */ <div
                                    style={{ fontSize: 9.5, color: "rgba(255,255,255,0.75)", marginTop: 2 }}
                                  >
                                    {role === "etudiant" ? masquerCinProf(cell.professeur) : cell.professeur}
                                  </div>
                                )}
                                <div
                                  style={{
                                    display: "inline-block",
                                    fontSize: 8.5,
                                    color: "#ffffff",
                                    background: "rgba(255,255,255,0.2)",
                                    fontWeight: 700,
                                    marginTop: 5,
                                    padding: "1px 6px",
                                    borderRadius: 10,
                                  }}
                                >
                                  {cell.estRattrapage
                                    ? "🔁 Rattrapage"
                                    : cell.typeRecurrence === "Unique"
                                      ? "◆ Occurrence unique"
                                      : "↻ Hebdomadaire"}
                                </div>
                                {(cell.lienTeams || cell.lienEnregistrements) && (
                                  /* @__PURE__ */ <div
                                    style={{ display: "flex", flexWrap: "wrap", gap: 4, marginTop: 5 }}
                                  >
                                    {cell.lienTeams && (
                                      /* @__PURE__ */ <a
                                        href={cell.lienTeams}
                                        target="_blank"
                                        rel="noopener noreferrer"
                                        onClick={(e) => e.stopPropagation()}
                                        style={{
                                          display: "inline-flex",
                                          alignItems: "center",
                                          gap: 3,
                                          fontSize: 9.5,
                                          fontWeight: 700,
                                          color: "#232733",
                                          background: "rgba(255,255,255,0.92)",
                                          border: "1px solid rgba(255,255,255,0.4)",
                                          borderRadius: 12,
                                          padding: "2px 7px",
                                          textDecoration: "none",
                                          boxShadow: "0 1px 3px rgba(0,0,0,0.25)",
                                        }}
                                      >
                                        💻 Rejoindre Teams
                                      </a>
                                    )}
                                    {cell.lienEnregistrements && (
                                      /* @__PURE__ */ <a
                                        href={cell.lienEnregistrements}
                                        target="_blank"
                                        rel="noopener noreferrer"
                                        onClick={(e) => e.stopPropagation()}
                                        style={{
                                          display: "inline-flex",
                                          alignItems: "center",
                                          gap: 3,
                                          fontSize: 9.5,
                                          fontWeight: 700,
                                          color: "#232733",
                                          background: "rgba(255,255,255,0.92)",
                                          border: "1px solid rgba(255,255,255,0.4)",
                                          borderRadius: 12,
                                          padding: "2px 7px",
                                          textDecoration: "none",
                                          boxShadow: "0 1px 3px rgba(0,0,0,0.25)",
                                        }}
                                      >
                                        🎥 Enregistrements
                                      </a>
                                    )}
                                  </div>
                                )}
                                {canEdit && (
                                  /* @__PURE__ */ <button
                                    onClick={(e) => {
                                      e.stopPropagation();
                                      requestDelete(cell, jour);
                                    }}
                                    onMouseEnter={(e) => {
                                      e.currentTarget.style.background = C.red;
                                      e.currentTarget.style.color = "white";
                                    }}
                                    onMouseLeave={(e) => {
                                      e.currentTarget.style.background = "white";
                                      e.currentTarget.style.color = C.red;
                                    }}
                                    style={{
                                      position: "absolute",
                                      top: -6,
                                      right: -6,
                                      width: 18,
                                      height: 18,
                                      display: "flex",
                                      alignItems: "center",
                                      justifyContent: "center",
                                      border: `1px solid ${C.red}55`,
                                      borderRadius: "50%",
                                      background: "white",
                                      color: C.red,
                                      cursor: "pointer",
                                      fontSize: 10,
                                      lineHeight: 1,
                                      boxShadow: "0 1px 4px rgba(0,0,0,0.15)",
                                      transition: "background 0.12s, color 0.12s",
                                    }}
                                  >
                                    ✕
                                  </button>
                                )}
                              </div>
                            ))}
                            {canEdit ? (
                              /* @__PURE__ */ <button
                                onClick={() => openAdd(jour, seance)}
                                onMouseEnter={(e) => {
                                  e.currentTarget.style.background = `linear-gradient(135deg, ${C.navyLight} 0%, #dce8f7 100%)`;
                                  e.currentTarget.style.borderColor = C.navy + "70";
                                  e.currentTarget.style.color = C.navy;
                                  e.currentTarget.style.boxShadow = "0 2px 8px rgba(26,47,94,0.15)";
                                }}
                                onMouseLeave={(e) => {
                                  e.currentTarget.style.background =
                                    "linear-gradient(135deg, #f7f9fc 0%, #eef1f6 100%)";
                                  e.currentTarget.style.borderColor = C.border;
                                  e.currentTarget.style.color = C.muted;
                                  e.currentTarget.style.boxShadow = "none";
                                }}
                                style={{
                                  width: "100%",
                                  background: "linear-gradient(135deg, #f7f9fc 0%, #eef1f6 100%)",
                                  border: `1.5px dashed ${C.border}`,
                                  borderRadius: 10,
                                  padding: cellItems.length ? "5px 8px" : "13px 8px",
                                  cursor: "pointer",
                                  color: C.muted,
                                  fontSize: cellItems.length ? 11.5 : 17,
                                  fontWeight: cellItems.length ? 600 : 400,
                                  transition: "background 0.15s, border-color 0.15s, color 0.15s, box-shadow 0.15s",
                                }}
                                title={
                                  cellItems.length ? "Ajouter un autre module sur ce créneau" : "Ajouter un module"
                                }
                              >
                                {cellItems.length ? "+ autre module" : "+"}
                              </button>
                            ) : cellItems.length === 0 ? (
                              /* @__PURE__ */ <span style={{ color: C.border, fontSize: 11 }}>—</span>
                            ) : null}
                          </div>
                        </td>
                      );
                    })}
                  </tr>
                ))}
              </tbody>
            </table>
          </Card>
        </div>
      )}
      {filterFiliere && modulesDuGroupe.length > 0 && (
        /* @__PURE__ */ <Card
          style={{ overflow: "hidden", marginTop: 20, background: "linear-gradient(160deg, #eef4fb 0%, #ffffff 55%)" }}
        >
          <div style={{ padding: "14px 18px 0" }}>
            <h4 style={{ margin: "0 0 4px", color: C.navy, fontSize: 14.5 }}>📅 Contrôle continu & Examens</h4>
            <div style={{ fontSize: 11.5, color: C.muted, marginBottom: 10 }}>
              {canEdit
                ? "Modifier une date la synchronise automatiquement avec le Calendrier scolaire."
                : "Dates communiquées par l'administration."}
            </div>
          </div>
          <DataTable
            cols={[
              { key: "intitule", label: "Module" },
              { key: "professeur", label: "Professeur", render: (v) => (role === "etudiant" ? masquerCinProf(v) : v) },
              {
                key: "_cc",
                label: "Contrôle continu",
                render: (_, m) => {
                  const cd = joursRestants(m.dateCC);
                  return canEdit ? (
                    /* @__PURE__ */ <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                      <input
                        type="date"
                        value={m.dateCC || ""}
                        onChange={(e) => handleSetDate(m, "dateCC", e.target.value)}
                        style={{
                          padding: "6px 10px",
                          border: `1px solid ${C.border}`,
                          borderRadius: 8,
                          fontSize: 12,
                          fontFamily: "inherit",
                        }}
                      />
                      {cd && (
                        /* @__PURE__ */ <span style={{ fontSize: 10.5, fontWeight: 700, color: cd.color }}>
                          {cd.texte}
                        </span>
                      )}
                    </div>
                  ) : (
                    /* @__PURE__ */ <span>
                      {m.dateCC ? formatDateFR(m.dateCC) : "—"} {cd && (
                        /* @__PURE__ */ <span
                          style={{ fontSize: 10.5, fontWeight: 700, color: cd.color, marginLeft: 6 }}
                        >
                          {cd.texte}
                        </span>
                      )}
                    </span>
                  );
                },
              },
              {
                key: "_examen",
                label: "Examen de fin de module",
                render: (_, m) => {
                  const cd = joursRestants(m.dateExamen);
                  return canEdit ? (
                    /* @__PURE__ */ <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                      <input
                        type="date"
                        value={m.dateExamen || ""}
                        onChange={(e) => handleSetDate(m, "dateExamen", e.target.value)}
                        style={{
                          padding: "6px 10px",
                          border: `1px solid ${C.border}`,
                          borderRadius: 8,
                          fontSize: 12,
                          fontFamily: "inherit",
                        }}
                      />
                      {cd && (
                        /* @__PURE__ */ <span style={{ fontSize: 10.5, fontWeight: 700, color: cd.color }}>
                          {cd.texte}
                        </span>
                      )}
                    </div>
                  ) : (
                    /* @__PURE__ */ <span>
                      {m.dateExamen ? formatDateFR(m.dateExamen) : "—"} {cd && (
                        /* @__PURE__ */ <span
                          style={{ fontSize: 10.5, fontWeight: 700, color: cd.color, marginLeft: 6 }}
                        >
                          {cd.texte}
                        </span>
                      )}
                    </span>
                  );
                },
              },
            ]}
            rows={modulesDuGroupe}
            emptyMsg="Aucun module trouvé"
            compact={true}
          />
        </Card>
      )}
      {showModal && (
        /* @__PURE__ */ <Modal
          title={editItem ? "Modifier le créneau" : "Ajouter un créneau → SharePoint"}
          onClose={() => setShowModal(false)}
        >
          <div className="eco-grid2" style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "0 14px" }}>
            <Sel
              label="Filière *"
              value={form.filiere}
              onChange={(e) => setForm((f) => ({ ...f, filiere: e.target.value, module: "" }))}
            >
              <option value="">-- Sélectionner --</option>
              {filieresOptions.map((f) => (
                /* @__PURE__ */ <option key={f} value={f}>
                  {f}
                </option>
              ))}
            </Sel>
            <CreatableSelect
              label="Semestre *"
              value={form.semestre}
              onChange={(v) => setForm((f) => ({ ...f, semestre: v, module: "" }))}
              options={semestresDisponibles}
            />
          </div>
          <div className="eco-grid2" style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "0 14px" }}>
            <Sel
              label="Jour *"
              value={form.jour}
              onChange={(e) =>
                setForm((f) => ({
                  ...f,
                  jour: e.target.value,
                  dateDebut: e.target.value ? dateDuJour(e.target.value) : f.dateDebut,
                }))
              }
            >
              <option value="">-- Sélectionner --</option>
              {(() => {
                const fi = (data.filieres || []).find((f) => f.intitule === form.filiere);
                const type = fi?.typeSeances || "Soir";
                const jours = estSemaineRamadan
                  ? type === "Jour"
                    ? JOURS_GROUPE_JOUR
                    : JOURS_RAMADAN_SOIR
                  : type === "Jour"
                    ? JOURS_GROUPE_JOUR
                    : JOURS_GROUPE_SOIR;
                return jours.map((j) => /* @__PURE__ */ <option key={j}>{j}</option>);
              })()}
            </Sel>
            <Sel
              label="Séance *"
              value={form.seance}
              onChange={(e) => setForm((f) => ({ ...f, seance: e.target.value }))}
            >
              <option value="">-- Sélectionner --</option>
              {(() => {
                const fi = (data.filieres || []).find((f) => f.intitule === form.filiere);
                const type = fi?.typeSeances || "Soir";
                const options = estSemaineRamadan
                  ? type === "Jour"
                    ? SEANCES_RAMADAN_JOUR
                    : SEANCES_RAMADAN_SOIR
                  : type === "Jour"
                    ? SEANCES_HORAIRE.filter((s) => s !== "19h00 - 21h00")
                    : SEANCES_HORAIRE.filter((s) => s === "19h00 - 21h00");
                return options.map((s) => /* @__PURE__ */ <option key={s}>{s}</option>);
              })()}
            </Sel>
          </div>
          <Sel
            label="Module *"
            value={form.module}
            onChange={(e) => setForm((f) => ({ ...f, module: e.target.value }))}
          >
            <option value="">-- Sélectionner --</option>
            {modulesOptions.map((c) => (
              /* @__PURE__ */ <option key={c.id} value={c.intitule}>
                {c.intitule}
              </option>
            ))}
          </Sel>
          {moduleChoisi && (
            /* @__PURE__ */ <div
              style={{
                background: C.light,
                borderRadius: 8,
                padding: "8px 12px",
                fontSize: 12,
                color: "#555",
                marginBottom: 14,
              }}
            >
              Professeur : <strong>{moduleChoisi.professeur || "—"}</strong>
              {moduleChoisi.salle ? ` · Salle ${moduleChoisi.salle}` : ""}
            </div>
          )}
          <label
            style={{
              display: "flex",
              alignItems: "center",
              gap: 9,
              padding: "10px 12px",
              background: "#f6f0fb",
              border: "1px solid #d9c3ec",
              borderRadius: 8,
              marginBottom: 14,
              cursor: "pointer",
            }}
          >
            <input
              type="checkbox"
              checked={form.estRattrapage}
              onChange={(e) =>
                setForm((f) => ({
                  ...f,
                  estRattrapage: e.target.checked,
                  typeRecurrence: e.target.checked ? "Unique" : f.typeRecurrence,
                }))
              }
              style={{ width: 16, height: 16, cursor: "pointer" }}
            />
            <span style={{ fontSize: 12.5, fontWeight: 700, color: "#7c3aed" }}>
              🔁 Ceci est une séance de rattrapage (jamais périodique)
            </span>
          </label>
          <Sel
            label="Type de récurrence *"
            value={form.estRattrapage ? "Unique" : form.typeRecurrence}
            disabled={form.estRattrapage}
            onChange={(e) => setForm((f) => ({ ...f, typeRecurrence: e.target.value }))}
          >
            <option value="Hebdomadaire">↻ Hebdomadaire (se répète chaque semaine)</option>
            <option value="Unique">◆ Occurrence unique (une seule séance)</option>
          </Sel>
          <div
            style={{
              display: "grid",
              gridTemplateColumns: form.typeRecurrence === "Hebdomadaire" && !form.estRattrapage ? "1fr 1fr" : "1fr",
              gap: "0 14px",
            }}
          >
            <Input
              label={
                form.typeRecurrence === "Hebdomadaire" && !form.estRattrapage
                  ? "Débute la semaine du"
                  : "Date de la séance"
              }
              type="date"
              value={form.dateDebut || ""}
              onChange={(e) => setForm((f) => ({ ...f, dateDebut: e.target.value }))}
            />
            {form.typeRecurrence === "Hebdomadaire" && !form.estRattrapage && (
              /* @__PURE__ */ <Input
                label="Se termine le"
                type="date"
                value={form.dateFin || ""}
                onChange={(e) => setForm((f) => ({ ...f, dateFin: e.target.value }))}
              />
            )}
          </div>
          {form.typeRecurrence === "Hebdomadaire" && !form.estRattrapage && (
            /* @__PURE__ */ <div style={{ fontSize: 11, color: C.muted, marginTop: -10, marginBottom: 14 }}>
              Par défaut, la série s'arrête à la fin du semestre (Paramétrage) — modifiable si besoin.
            </div>
          )}
          <Input
            label="🔗 Lien Teams (optionnel — si cours en ligne)"
            type="url"
            value={form.lienTeams || ""}
            onChange={(e) => setForm((f) => ({ ...f, lienTeams: e.target.value }))}
            placeholder="https://teams.microsoft.com/l/meetup-join/..."
          />
          <Input
            label="🎥 Lien du dossier d'enregistrements (optionnel)"
            type="url"
            value={form.lienEnregistrements || ""}
            onChange={(e) => setForm((f) => ({ ...f, lienEnregistrements: e.target.value }))}
            placeholder="https://ecolesuperieurederabat.sharepoint.com/sites/.../Recordings"
          />
          <div style={{ display: "flex", gap: 10, justifyContent: "flex-end" }}>
            <Btn variant="light" onClick={() => setShowModal(false)}>
              Annuler
            </Btn>
            <Btn
              onClick={handleSave}
              disabled={
                !form.filiere || !form.semestre || !form.jour || !form.seance || !form.module || !form.dateDebut
              }
            >
              {editItem ? "Mettre à jour → SharePoint" : "Enregistrer → SharePoint"}
            </Btn>
          </div>
        </Modal>
      )}
      {showDeleteChoice && (
        /* @__PURE__ */ <Modal
          title="Supprimer ce créneau"
          onClose={() => setShowDeleteChoice(null)}
          hideFooterClose={true}
        >
          <p style={{ fontSize: 13, color: "#444", lineHeight: 1.6, marginBottom: 18 }}>
            Le créneau <strong>"{showDeleteChoice.item.module}"</strong> ({showDeleteChoice.item.jour} · {showDeleteChoice.item.seance}) se répète chaque semaine. Que souhaitez-vous faire ?
          </p>
          <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
            <Btn
              onClick={() => {
                deleteOccurrenceSeule(showDeleteChoice.item, showDeleteChoice.jour);
                setShowDeleteChoice(null);
              }}
            >
              Annuler uniquement cette semaine ({formatDateFR(dateDuJour(showDeleteChoice.jour))})
            </Btn>
            <Btn
              variant="danger"
              onClick={() => {
                deleteSerieEntiere(showDeleteChoice.item);
                setShowDeleteChoice(null);
              }}
            >
              🗑 Supprimer toute la série
            </Btn>
            <Btn variant="light" onClick={() => setShowDeleteChoice(null)}>
              Ne rien faire
            </Btn>
          </div>
        </Modal>
      )}
    </div>
  );
}
function waLink(phone, message) {
  const chiffres = (phone || "").replace(/\D/g, "");
  if (!chiffres) return null;
  const intl = chiffres.startsWith("212")
    ? chiffres
    : chiffres.startsWith("0")
      ? "212" + chiffres.slice(1)
      : "212" + chiffres;
  return `https://wa.me/${intl}?text=${encodeURIComponent(message)}`;
}
function Annonces({ data, setData, role, user, toast }) {
  const canEdit = ["administrateur", "directrice"].includes(role);
  const [showModal, setShowModal] = useState(false);
  const [showWhatsApp, setShowWhatsApp] = useState(null);
  const [form, setForm] = useState({
    titre: "",
    contenu: "",
    auteur: "",
    cible: "Tous",
    date: /* @__PURE__ */ new Date().toISOString().split("T")[0],
  });
  const handleSave = () => {
    const newA = { ...form, id: Date.now() };
    const spF = {
      Title: form.titre || "",
      Titre: form.titre || "",
      Contenu: form.contenu || "",
      Auteur: form.auteur || "",
      Cible: form.cible || "Tous",
      DateAnnonce: form.date || "",
    };
    if (!user?.token) {
      toast("⚠️ Non connecté à SharePoint — annonce non publiée");
    } else {
      saveToSP("ECOGEST_Annonces", spF, user.token).then((result) => {
        if (!result) {
          toast("⚠️ Échec de la publication dans SharePoint — annonce non ajoutée");
          return;
        }
        setData((d) => ({ ...d, annonces: [{ ...newA, _spId: result.id }, ...d.annonces] }));
        toast("Annonce publiée dans SharePoint ✓");
      });
    }
    setShowModal(false);
    setForm({
      titre: "",
      contenu: "",
      auteur: "",
      cible: "Tous",
      date: /* @__PURE__ */ new Date().toISOString().split("T")[0],
    });
  };
  return (
    /* @__PURE__ */ <div>
      <PageHeader
        title="Annonces"
        sub={`${data.annonces.length} annonce(s) publiée(s) sur ECOGEST`}
        action={canEdit && /* @__PURE__ */ <Btn onClick={() => setShowModal(true)}>+ Publier une annonce</Btn>}
      />
      <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
        {data.annonces.map((a) => (
          /* @__PURE__ */ <Card key={a.id} style={{ padding: 20 }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
              <div style={{ display: "flex", gap: 14 }}>
                <div
                  style={{
                    width: 40,
                    height: 40,
                    borderRadius: 10,
                    background: C.navyLight,
                    display: "flex",
                    alignItems: "center",
                    justifyContent: "center",
                    fontSize: 20,
                    flexShrink: 0,
                  }}
                >
                  📢
                </div>
                <div>
                  <h4 style={{ margin: "0 0 6px", color: C.navy, fontSize: 15 }}>{a.titre}</h4>
                  <p style={{ margin: "0 0 8px", fontSize: 13, color: "#444", lineHeight: 1.6 }}>{a.contenu}</p>
                  <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
                    <span style={{ fontSize: 12, color: C.muted }}>
                      {a.auteur} · {formatDateFR(a.date)}
                    </span>
                    <Badge color={C.navy} size={11}>
                      {a.cible}
                    </Badge>
                  </div>
                </div>
              </div>
              <button
                onClick={() => setShowWhatsApp(a)}
                title="Envoyer via WhatsApp"
                style={{
                  display: "flex",
                  alignItems: "center",
                  gap: 6,
                  border: "none",
                  background: "#25D366",
                  color: "white",
                  fontSize: 12,
                  fontWeight: 700,
                  padding: "7px 14px",
                  borderRadius: 20,
                  cursor: "pointer",
                  fontFamily: "inherit",
                  flexShrink: 0,
                  boxShadow: "0 2px 8px rgba(37,211,102,0.35)",
                }}
              >
                📱 WhatsApp
              </button>
            </div>
          </Card>
        ))}
      </div>
      {showModal && (
        /* @__PURE__ */ <Modal title="Publier une annonce → SharePoint" onClose={() => setShowModal(false)}>
          <Input
            label="Titre *"
            value={form.titre}
            onChange={(e) => setForm((f) => ({ ...f, titre: e.target.value }))}
          />
          <div style={{ marginBottom: 13 }}>
            <label style={{ fontSize: 12, fontWeight: 600, color: "#555", display: "block", marginBottom: 4 }}>
              Contenu *
            </label>
            <textarea
              value={form.contenu}
              onChange={(e) => setForm((f) => ({ ...f, contenu: e.target.value }))}
              rows={4}
              style={{
                width: "100%",
                padding: "8px 11px",
                border: `1px solid ${C.border}`,
                borderRadius: 8,
                fontSize: 13,
                fontFamily: "inherit",
                resize: "vertical",
                boxSizing: "border-box",
              }}
            />
          </div>
          <div className="eco-grid3" style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: "0 14px" }}>
            <Input
              label="Auteur"
              value={form.auteur}
              onChange={(e) => setForm((f) => ({ ...f, auteur: e.target.value }))}
            />
            <Sel
              label="Destinataires"
              value={form.cible}
              onChange={(e) => setForm((f) => ({ ...f, cible: e.target.value }))}
            >
              <option>Tous</option>
              <option>Étudiants</option>
              <option>Professeurs</option>
              <option>Administration</option>
            </Sel>
            <Input
              label="Date"
              type="date"
              value={form.date}
              onChange={(e) => setForm((f) => ({ ...f, date: e.target.value }))}
            />
          </div>
          <div style={{ display: "flex", gap: 10, justifyContent: "flex-end" }}>
            <Btn variant="light" onClick={() => setShowModal(false)}>
              Annuler
            </Btn>
            <Btn onClick={handleSave} disabled={!form.titre || !form.contenu}>
              Publier → SharePoint
            </Btn>
          </div>
        </Modal>
      )}
      {showWhatsApp &&
        (() => {
          const message = `📢 ${showWhatsApp.titre}

${showWhatsApp.contenu}`;
          const destinataires = [
            ...(["Tous", "Étudiants"].includes(showWhatsApp.cible)
              ? (data.etudiants || []).map((e) => ({
                  nom: `${e.prenom} ${e.nom}`,
                  telephone: e.telephone,
                  type: "Étudiant",
                }))
              : []),
            ...(["Tous", "Professeurs"].includes(showWhatsApp.cible)
              ? (data.professeurs || []).map((p) => ({
                  nom: `${p.prenom} ${p.nom}`,
                  telephone: p.telephone,
                  type: "Professeur",
                }))
              : []),
          ];
          return (
            /* @__PURE__ */ <Modal
              title={`Envoyer "${showWhatsApp.titre}" via WhatsApp`}
              onClose={() => setShowWhatsApp(null)}
              width={520}
            >
              <div
                style={{
                  background: "#e9f9ef",
                  border: "1px solid #b8ecc9",
                  borderRadius: 8,
                  padding: "10px 14px",
                  fontSize: 12,
                  color: "#1a7a3d",
                  marginBottom: 16,
                  lineHeight: 1.6,
                }}
              >{`📱 Un clic ouvre WhatsApp avec le message déjà rempli pour ce contact — il suffit ensuite d'appuyer sur "Envoyer" dans WhatsApp. Ce n'est pas un envoi groupé automatique : un clic = un envoi, contact par contact.`}</div>
              {destinataires.length === 0 ? (
                /* @__PURE__ */ <div style={{ textAlign: "center", padding: 24, color: C.muted, fontSize: 13 }}>
                  Aucun contact avec numéro de téléphone pour cette cible ("{showWhatsApp.cible}").
                </div>
              ) : (
                destinataires.map((d, i) => {
                  const lien = waLink(d.telephone, message);
                  return (
                    /* @__PURE__ */ <div
                      key={i}
                      style={{
                        display: "flex",
                        justifyContent: "space-between",
                        alignItems: "center",
                        padding: "9px 4px",
                        borderBottom: `1px solid ${C.light}`,
                      }}
                    >
                      <div>
                        <div style={{ fontWeight: 600, fontSize: 13 }}>{d.nom}</div>
                        <div style={{ fontSize: 11, color: C.muted }}>
                          {d.type} · {d.telephone || "—"}
                        </div>
                      </div>
                      {lien ? (
                        /* @__PURE__ */ <a
                          href={lien}
                          target="_blank"
                          rel="noopener noreferrer"
                          style={{
                            display: "inline-flex",
                            alignItems: "center",
                            gap: 5,
                            background: "#25D366",
                            color: "white",
                            fontSize: 11.5,
                            fontWeight: 700,
                            padding: "6px 12px",
                            borderRadius: 18,
                            textDecoration: "none",
                          }}
                        >
                          📱 Envoyer
                        </a>
                      ) : (
                        /* @__PURE__ */ <span style={{ fontSize: 11, color: C.muted, fontStyle: "italic" }}>
                          Pas de téléphone
                        </span>
                      )}
                    </div>
                  );
                })
              )}
            </Modal>
          );
        })()}
    </div>
  );
}
const TYPES_EVENEMENT = {
  Examen: { color: C.red, icon: "📝", bgGradient: "linear-gradient(135deg, #fdeeec 0%, #fad0ca 100%)" },
  Vacances: { color: C.green, icon: "🏖️", bgGradient: "linear-gradient(135deg, #eafaf0 0%, #c7ecd8 100%)" },
  Rentrée: { color: C.navy, icon: "🎒", bgGradient: "linear-gradient(135deg, #eef4fb 0%, #cddcf3 100%)" },
  Réunion: { color: C.purple, icon: "🗓️", bgGradient: "linear-gradient(135deg, #f6f0fb 0%, #e2ccf7 100%)" },
  Autre: { color: C.orange, icon: "📌", bgGradient: "linear-gradient(135deg, #fff3e6 0%, #fbdcb3 100%)" },
};
function Calendrier({ data, setData, role, user, toast }) {
  const canEdit = ["administrateur", "directrice"].includes(role);
  const [showModal, setShowModal] = useState(false);
  const [editItem, setEditItem] = useState(null);
  const [filterType, setFilterType] = useState("");
  const emptyForm = { titre: "", type: "Examen", dateDebut: today(), dateFin: "", description: "", filiere: "" };
  const [form, setForm] = useState(emptyForm);
  const openAdd = () => {
    setEditItem(null);
    setForm(emptyForm);
    setShowModal(true);
  };
  const openEdit = (ev) => {
    setEditItem(ev);
    setForm({ ...ev });
    setShowModal(true);
  };
  const spFields = (f) => ({
    Title: f.titre || "",
    TypeEvenement: f.type || "Autre",
    DateDebut: f.dateDebut || "",
    DateFin: f.dateFin || "",
    Description: f.description || "",
    Filiere: f.filiere || "",
  });
  const handleSave = () => {
    if (!user?.token) {
      toast("⚠️ Non connecté à SharePoint — événement non enregistré");
      return;
    }
    if (editItem) {
      const newEv = { ...form, id: editItem.id, _spId: editItem._spId };
      setData((d) => ({ ...d, calendrier: (d.calendrier || []).map((e) => (e.id === editItem.id ? newEv : e)) }));
      saveToSP("ECOGEST_Calendrier", spFields(form), user.token, editItem._spId).then((result) => {
        if (!result) {
          toast("⚠️ Échec de la mise à jour dans SharePoint");
          return;
        }
        toast("Événement mis à jour dans SharePoint ✓");
      });
    } else {
      saveToSP("ECOGEST_Calendrier", spFields(form), user.token).then((result) => {
        if (!result) {
          toast("⚠️ Échec de l'enregistrement dans SharePoint — événement non ajouté");
          return;
        }
        setData((d) => ({
          ...d,
          calendrier: [...(d.calendrier || []), { ...form, id: Date.now(), _spId: result.id }],
        }));
        toast("Événement ajouté dans SharePoint ✓");
      });
    }
    setShowModal(false);
  };
  const handleDelete = (ev) => {
    if (!window.confirm(`Supprimer l'événement "${ev.titre}" ?`)) return;
    setData((d) => ({ ...d, calendrier: (d.calendrier || []).filter((e) => e.id !== ev.id) }));
    if (user?.token && ev._spId) deleteFromSP("ECOGEST_Calendrier", ev._spId, user.token);
    toast("Événement supprimé de SharePoint ✓");
  };
  const evenements = (data.calendrier || [])
    .filter((e) => !filterType || e.type === filterType)
    .sort((a, b) => (a.dateDebut || "").localeCompare(b.dateDebut || ""));
  const aVenir = (data.calendrier || []).filter((e) => e.dateDebut >= today());
  const examensAVenir = aVenir.filter((e) => e.type === "Examen").length;
  const vacancesAVenir = aVenir.filter((e) => e.type === "Vacances").length;
  const parMois = {};
  evenements.forEach((e) => {
    const mois = (e.dateDebut || "").slice(0, 7) || "Sans date";
    if (!parMois[mois]) parMois[mois] = [];
    parMois[mois].push(e);
  });
  const moisLabel = (m) => {
    if (m === "Sans date") return m;
    const [y, mo] = m.split("-");
    const noms = [
      "Janvier",
      "Février",
      "Mars",
      "Avril",
      "Mai",
      "Juin",
      "Juillet",
      "Août",
      "Septembre",
      "Octobre",
      "Novembre",
      "Décembre",
    ];
    return `${noms[parseInt(mo, 10) - 1]} ${y}`;
  };
  return (
    /* @__PURE__ */ <div>
      <PageHeader
        title="Calendrier scolaire"
        sub={`${(data.calendrier || []).length} événement(s) — liste SharePoint "ECOGEST_Calendrier"`}
        action={canEdit && /* @__PURE__ */ <Btn onClick={openAdd}>+ Nouvel événement</Btn>}
      />
      <div style={{ display: "flex", gap: 12, marginBottom: 20 }}>
        <StatCard label="Total événements" value={(data.calendrier || []).length} color={C.navy} icon="🗓️" />
        <StatCard label="Examens à venir" value={examensAVenir} color={C.red} icon="📝" />
        <StatCard label="Vacances à venir" value={vacancesAVenir} color={C.green} icon="🏖️" />
      </div>
      <div style={{ display: "flex", gap: 8, marginBottom: 18, flexWrap: "wrap" }}>
        {["", ...Object.keys(TYPES_EVENEMENT)].map((t) => {
          const col = t ? TYPES_EVENEMENT[t].color : C.navy;
          const actif = filterType === t;
          return (
            /* @__PURE__ */ <button
              key={t}
              onClick={() => setFilterType(t)}
              style={{
                padding: "7px 16px",
                borderRadius: 20,
                border: `1px solid ${actif ? col : col + "45"}`,
                background: actif ? col : col + "18",
                color: actif ? C.white : col,
                fontSize: 12.5,
                cursor: "pointer",
                fontFamily: "inherit",
                fontWeight: 700,
                boxShadow: actif ? `0 3px 10px ${col}55` : "none",
                transition: "all 0.15s",
              }}
            >
              {t ? `${TYPES_EVENEMENT[t].icon} ${t}` : "🗂️ Tous"}
            </button>
          );
        })}
      </div>
      {Object.keys(parMois).length === 0 ? (
        /* @__PURE__ */ <Card style={{ padding: 32, textAlign: "center", color: C.muted, fontSize: 13 }}>
          Aucun événement au calendrier
        </Card>
      ) : (
        Object.keys(parMois)
          .sort()
          .map((mois) => (
            /* @__PURE__ */ <div key={mois} style={{ marginBottom: 20 }}>
              <div
                style={{
                  fontSize: 12,
                  fontWeight: 800,
                  color: C.navy,
                  textTransform: "uppercase",
                  letterSpacing: "0.06em",
                  marginBottom: 8,
                }}
              >
                {moisLabel(mois)}
              </div>
              <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
                {parMois[mois].map((ev) => {
                  const t = TYPES_EVENEMENT[ev.type] || TYPES_EVENEMENT.Autre;
                  return (
                    /* @__PURE__ */ <Card
                      key={ev.id}
                      style={{
                        padding: "14px 18px",
                        display: "flex",
                        justifyContent: "space-between",
                        alignItems: "center",
                        gap: 14,
                        borderLeft: `4px solid ${t.color}`,
                        background: t.bgGradient,
                        boxShadow: "0 3px 10px rgba(20,30,60,0.1)",
                      }}
                    >
                      <div style={{ display: "flex", gap: 14, alignItems: "center", minWidth: 0 }}>
                        <div
                          style={{
                            width: 38,
                            height: 38,
                            borderRadius: 10,
                            background: "rgba(255,255,255,0.75)",
                            boxShadow: `0 2px 6px ${t.color}40`,
                            display: "flex",
                            alignItems: "center",
                            justifyContent: "center",
                            fontSize: 18,
                            flexShrink: 0,
                          }}
                        >
                          {t.icon}
                        </div>
                        <div style={{ minWidth: 0 }}>
                          <div style={{ fontWeight: 700, fontSize: 14, color: "#232733" }}>{ev.titre}</div>
                          <div style={{ fontSize: 12, color: "#5a6172", marginTop: 2 }}>
                            {formatDateFR(ev.dateDebut)}
                            {ev.dateFin && ev.dateFin !== ev.dateDebut ? ` → ${formatDateFR(ev.dateFin)}` : ""}
                            {ev.filiere ? ` · ${ev.filiere}` : ""}
                          </div>
                          {ev.description && (
                            /* @__PURE__ */ <div style={{ fontSize: 12.5, color: "#4a5164", marginTop: 4 }}>
                              {ev.description}
                            </div>
                          )}
                        </div>
                      </div>
                      <div style={{ display: "flex", gap: 8, alignItems: "center", flexShrink: 0 }}>
                        <span
                          style={{
                            background: t.color,
                            color: "#ffffff",
                            fontSize: 11.5,
                            fontWeight: 700,
                            padding: "4px 11px",
                            borderRadius: 20,
                            boxShadow: `0 2px 6px ${t.color}55`,
                          }}
                        >
                          {ev.type}
                        </span>
                        {canEdit && (
                          /* @__PURE__ */ <React.Fragment>
                            <button
                              title="Modifier"
                              onClick={() => openEdit(ev)}
                              style={{
                                border: "none",
                                borderRadius: 8,
                                cursor: "pointer",
                                width: 30,
                                height: 30,
                                fontSize: 13,
                                background: "rgba(255,255,255,0.85)",
                                color: C.navy,
                                boxShadow: "0 1px 3px rgba(20,30,60,0.15)",
                              }}
                            >
                              ✏️
                            </button>
                            <button
                              title="Supprimer"
                              onClick={() => handleDelete(ev)}
                              style={{
                                border: "none",
                                borderRadius: 8,
                                cursor: "pointer",
                                width: 30,
                                height: 30,
                                fontSize: 13,
                                background: "rgba(255,255,255,0.85)",
                                color: C.red,
                                boxShadow: "0 1px 3px rgba(20,30,60,0.15)",
                              }}
                            >
                              🗑
                            </button>
                          </React.Fragment>
                        )}
                      </div>
                    </Card>
                  );
                })}
              </div>
            </div>
          ))
      )}
      {showModal && (
        /* @__PURE__ */ <Modal
          title={editItem ? "Modifier l'événement" : "Nouvel événement → SharePoint"}
          onClose={() => setShowModal(false)}
        >
          <Input
            label="Titre *"
            value={form.titre}
            onChange={(e) => setForm((f) => ({ ...f, titre: e.target.value }))}
            placeholder="Ex: Examens de rattrapage S1"
          />
          <Sel label="Type *" value={form.type} onChange={(e) => setForm((f) => ({ ...f, type: e.target.value }))}>
            {Object.keys(TYPES_EVENEMENT).map((t) => (
              /* @__PURE__ */ <option key={t}>{t}</option>
            ))}
          </Sel>
          <div className="eco-grid2" style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "0 14px" }}>
            <Input
              label="Date de début *"
              type="date"
              value={form.dateDebut}
              onChange={(e) => setForm((f) => ({ ...f, dateDebut: e.target.value }))}
            />
            <Input
              label="Date de fin (si période)"
              type="date"
              value={form.dateFin}
              onChange={(e) => setForm((f) => ({ ...f, dateFin: e.target.value }))}
            />
          </div>
          <Sel
            label="Groupe concerné (optionnel)"
            value={form.filiere}
            onChange={(e) => setForm((f) => ({ ...f, filiere: e.target.value }))}
          >
            <option value="">-- Tous les groupes --</option>
            {getFilieres(data).map((f) => (
              /* @__PURE__ */ <option key={f} value={f}>
                {f}
              </option>
            ))}
          </Sel>
          <div style={{ marginBottom: 13 }}>
            <label style={{ fontSize: 12, fontWeight: 600, color: "#555", display: "block", marginBottom: 4 }}>
              Description
            </label>
            <textarea
              value={form.description}
              onChange={(e) => setForm((f) => ({ ...f, description: e.target.value }))}
              rows={3}
              style={{
                width: "100%",
                padding: "8px 11px",
                border: `1px solid ${C.border}`,
                borderRadius: 8,
                fontSize: 13,
                fontFamily: "inherit",
                resize: "vertical",
                boxSizing: "border-box",
              }}
            />
          </div>
          <div style={{ display: "flex", gap: 10, justifyContent: "flex-end" }}>
            <Btn variant="light" onClick={() => setShowModal(false)}>
              Annuler
            </Btn>
            <Btn onClick={handleSave} disabled={!form.titre || !form.dateDebut}>
              {editItem ? "Mettre à jour → SharePoint" : "Ajouter → SharePoint"}
            </Btn>
          </div>
        </Modal>
      )}
    </div>
  );
}
function Parametrage({ data, setData, user, toast }) {
  const [selected, setSelected] = useState(data.configuration?.anneeUniversitaireActive || ANNEES_UNIVERSITAIRES[0]);
  const [saving, setSaving] = useState(false);
  const handleSave = async () => {
    if (!user?.token) {
      toast("⚠️ Non connecté à SharePoint — modification non enregistrée");
      return;
    }
    setSaving(true);
    const spId = data.configuration?._spId;
    const result = await saveToSP(
      "ECOGEST_Parametres",
      { Title: "Configuration ECOGEST", AnneeUniversitaireActive: selected },
      user.token,
      spId
    );
    setSaving(false);
    if (!result) {
      toast("⚠️ Échec de l'enregistrement dans SharePoint");
      return;
    }
    setData((d) => ({
      ...d,
      configuration: { ...d.configuration, _spId: result.id, anneeUniversitaireActive: selected },
    }));
    toast("Année universitaire active mise à jour ✓");
  };
  const handleToggleFiliereTauxAbsence = (filiere) => {
    if (!user?.token) {
      toast("⚠️ Non connecté à SharePoint — modification non enregistrée");
      return;
    }
    const actuelles = data.configuration?.filieresAvecTauxAbsence || [];
    const nouvelles = actuelles.includes(filiere) ? actuelles.filter((f) => f !== filiere) : [...actuelles, filiere];
    const spId = data.configuration?._spId;
    setData((d) => ({ ...d, configuration: { ...d.configuration, filieresAvecTauxAbsence: nouvelles } }));
    saveToSP(
      "ECOGEST_Parametres",
      {
        Title: "Configuration ECOGEST",
        AnneeUniversitaireActive: data.configuration?.anneeUniversitaireActive || selected,
        RamadanDebut: data.configuration?.ramadanDebut || "",
        RamadanFin: data.configuration?.ramadanFin || "",
        FilieresAvecTauxAbsence: nouvelles.join(" | "),
      },
      user.token,
      spId
    ).then((result) => {
      if (!result) {
        toast("⚠️ Échec de l'enregistrement");
        return;
      }
      if (!spId) setData((d) => ({ ...d, configuration: { ...d.configuration, _spId: result.id } }));
    });
  };
  const handleSaveRamadan = (champ, valeur) => {
    if (!user?.token) {
      toast("⚠️ Non connecté à SharePoint — date non enregistrée");
      return;
    }
    const spId = data.configuration?._spId;
    const nouvelleConfig = { ...data.configuration, [champ]: valeur };
    setData((d) => ({ ...d, configuration: nouvelleConfig }));
    saveToSP(
      "ECOGEST_Parametres",
      {
        Title: "Configuration ECOGEST",
        AnneeUniversitaireActive: nouvelleConfig.anneeUniversitaireActive || selected,
        RamadanDebut: nouvelleConfig.ramadanDebut || "",
        RamadanFin: nouvelleConfig.ramadanFin || "",
      },
      user.token,
      spId
    ).then((result) => {
      if (!result) {
        toast("⚠️ Échec de l'enregistrement");
        return;
      }
      if (!spId) setData((d) => ({ ...d, configuration: { ...d.configuration, _spId: result.id } }));
    });
  };
  const semestresConnus = [
    .../* @__PURE__ */ new Set([
      ...SEMESTRES_MODULE_SEED,
      ...(data.cours || []).map((c) => c.semestre).filter(Boolean),
      ...(data.semestresDates || []).map((s) => s.semestre).filter(Boolean),
    ]),
  ].sort((a, b) => a.localeCompare(b, "fr", { numeric: true, sensitivity: "base" }));
  const handleSaveSemestreDate = (semestre, champ, valeur) => {
    if (!user?.token) {
      toast("⚠️ Non connecté à SharePoint — date non enregistrée");
      return;
    }
    const existant = (data.semestresDates || []).find((s) => s.semestre === semestre);
    if (existant) {
      const updated = { ...existant, [champ]: valeur };
      setData((d) => ({ ...d, semestresDates: d.semestresDates.map((s) => (s.semestre === semestre ? updated : s)) }));
      if (existant._spId && existant._spId !== "pending") {
        saveToSP(
          "ECOGEST_SemestresDates",
          { Semestre: semestre, DateDebut: updated.dateDebut || "", DateFin: updated.dateFin || "" },
          user.token,
          existant._spId
        ).then((result) => {
          if (!result) toast("⚠️ Échec de l'enregistrement");
        });
      }
      return;
    }
    const nouvelleEntreeLocale = {
      id: Date.now(),
      _spId: "pending",
      semestre,
      dateDebut: "",
      dateFin: "",
      [champ]: valeur,
    };
    setData((d) => ({ ...d, semestresDates: [...d.semestresDates, nouvelleEntreeLocale] }));
    saveToSP(
      "ECOGEST_SemestresDates",
      {
        Title: semestre,
        Semestre: semestre,
        DateDebut: nouvelleEntreeLocale.dateDebut || "",
        DateFin: nouvelleEntreeLocale.dateFin || "",
      },
      user.token
    ).then((result) => {
      if (!result) {
        toast("⚠️ Échec de l'enregistrement");
        return;
      }
      setData((d) => {
        const actuel = d.semestresDates.find((s) => s.semestre === semestre && s._spId === "pending");
        const nouveaux = d.semestresDates.map((s) =>
          s.semestre === semestre && s._spId === "pending" ? { ...s, _spId: result.id } : s
        );
        if (actuel) {
          saveToSP(
            "ECOGEST_SemestresDates",
            { Semestre: semestre, DateDebut: actuel.dateDebut || "", DateFin: actuel.dateFin || "" },
            user.token,
            result.id
          );
        }
        return { ...d, semestresDates: nouveaux };
      });
    });
  };
  return (
    /* @__PURE__ */ <div>
      <PageHeader title="Paramétrage" sub="Réglages généraux appliqués à l'ensemble de l'application" />
      <Card style={{ padding: 24, maxWidth: 480, marginBottom: 20 }}>
        <div
          style={{
            fontSize: 11,
            fontWeight: 700,
            color: C.orange,
            textTransform: "uppercase",
            letterSpacing: "0.08em",
            marginBottom: 10,
          }}
        >
          Année universitaire active
        </div>
        <p style={{ fontSize: 13, color: "#555", marginBottom: 16, lineHeight: 1.6 }}>
          Cette année sera appliquée automatiquement à toutes les opérations de l'application (versements, rapports...)
          — plus besoin de la sélectionner manuellement à chaque fois.
        </p>
        <Sel value={selected} onChange={(e) => setSelected(e.target.value)}>
          {ANNEES_UNIVERSITAIRES.map((a) => (
            /* @__PURE__ */ <option key={a} value={a}>
              {a}
            </option>
          ))}
        </Sel>
        <div style={{ marginTop: 16 }}>
          <Btn onClick={handleSave} disabled={saving}>
            {saving ? "Enregistrement..." : "💾 Enregistrer"}
          </Btn>
        </div>
        <div
          style={{
            marginTop: 20,
            background: C.navyLight,
            borderRadius: 8,
            padding: "10px 14px",
            fontSize: 12,
            color: C.navy,
          }}
        >
          📌 Année actuellement active dans l'application : <strong>{data.configuration?.anneeUniversitaireActive || "—"}</strong>
        </div>
      </Card>
      <Card style={{ padding: 24, maxWidth: 620 }}>
        <div
          style={{
            fontSize: 11,
            fontWeight: 700,
            color: C.orange,
            textTransform: "uppercase",
            letterSpacing: "0.08em",
            marginBottom: 10,
          }}
        >
          Dates des semestres
        </div>
        <p style={{ fontSize: 13, color: "#555", marginBottom: 16, lineHeight: 1.6 }}>
          Utilisées pour générer les semaines sélectionnables dans l'Emploi du temps, et pour délimiter automatiquement
          les affectations hebdomadaires (elles s'arrêtent par défaut à la date de fin du semestre).
        </p>
        {semestresConnus.map((sem) => {
          const sd = (data.semestresDates || []).find((s) => s.semestre === sem) || {};
          return (
            /* @__PURE__ */ <div
              key={sem}
              style={{
                display: "grid",
                gridTemplateColumns: "1fr 1fr 1fr",
                gap: "0 14px",
                alignItems: "end",
                marginBottom: 4,
              }}
            >
              <div style={{ fontSize: 13, fontWeight: 700, color: C.navy, marginBottom: 13 }}>{sem}</div>
              <Input
                label="Début"
                type="date"
                value={sd.dateDebut || ""}
                onChange={(e) => handleSaveSemestreDate(sem, "dateDebut", e.target.value)}
              />
              <Input
                label="Fin"
                type="date"
                value={sd.dateFin || ""}
                onChange={(e) => handleSaveSemestreDate(sem, "dateFin", e.target.value)}
              />
            </div>
          );
        })}
      </Card>
      <Card style={{ padding: 24, maxWidth: 620, marginTop: 20 }}>
        <div
          style={{
            fontSize: 11,
            fontWeight: 700,
            color: C.orange,
            textTransform: "uppercase",
            letterSpacing: "0.08em",
            marginBottom: 10,
          }}
        >
          🌙 Période de Ramadan
        </div>
        <p style={{ fontSize: 13, color: "#555", marginBottom: 16, lineHeight: 1.6 }}>
          Pendant cette période, l'Emploi du temps applique automatiquement les horaires adaptés : groupes de journée →
          Lundi-Vendredi (12h00-13h30, 13h40-15h10) ; groupes du soir → Samedi-Dimanche (11h00-12h30, 13h00-14h30).
        </p>
        <div className="eco-grid2" style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "0 14px" }}>
          <Input
            label="Début du Ramadan"
            type="date"
            value={data.configuration?.ramadanDebut || ""}
            onChange={(e) => handleSaveRamadan("ramadanDebut", e.target.value)}
          />
          <Input
            label="Fin du Ramadan"
            type="date"
            value={data.configuration?.ramadanFin || ""}
            onChange={(e) => handleSaveRamadan("ramadanFin", e.target.value)}
          />
        </div>
      </Card>
      <Card style={{ padding: 24, maxWidth: 620, marginTop: 20 }}>
        <div
          style={{
            fontSize: 11,
            fontWeight: 700,
            color: C.orange,
            textTransform: "uppercase",
            letterSpacing: "0.08em",
            marginBottom: 10,
          }}
        >
          📉 Taux d'absence dans les relevés de notes
        </div>
        <p style={{ fontSize: 13, color: "#555", marginBottom: 16, lineHeight: 1.6 }}>
          Cochez les filières pour lesquelles le relevé de notes doit calculer et afficher le taux d'absence par module,
          et l'utiliser dans la décision du jury : <strong>le module est déclaré "Non validé" dès que ce taux atteint 25%</strong>, quelle que soit la note
          obtenue. Pour les filières non cochées, ce taux n'est jamais calculé ni affiché dans leurs relevés.
        </p>
        <div style={{ border: `1px solid ${C.border}`, borderRadius: 8, maxHeight: 220, overflowY: "auto" }}>
          {getFilieres(data).map((f, i, arr) => (
            /* @__PURE__ */ <label
              key={f}
              style={{
                display: "flex",
                alignItems: "center",
                gap: 9,
                padding: "9px 14px",
                borderBottom: i < arr.length - 1 ? `1px solid ${C.light}` : "none",
                cursor: "pointer",
                fontSize: 13,
              }}
            >
              <input
                type="checkbox"
                checked={(data.configuration?.filieresAvecTauxAbsence || []).includes(f)}
                onChange={() => handleToggleFiliereTauxAbsence(f)}
                style={{ width: 16, height: 16, cursor: "pointer" }}
              />
              {f}
            </label>
          ))}
        </div>
      </Card>
    </div>
  );
}
function Configuration({ toast, user }) {
  const [cfg, setCfg] = useState({ tenantId: SP.tenantId, clientId: SP.clientId, siteUrl: SP.siteUrl });
  const [saved, setSaved] = useState(false);
  const [testResult, setTestResult] = useState(null);
  const [testing, setTesting] = useState(false);
  const handleTest = async () => {
    setTesting(true);
    setTestResult(null);
    try {
      if (!user?.token) {
        setTestResult({ ok: false, msg: "Non connecté — reconnectez-vous." });
        setTesting(false);
        return;
      }
      const siteId = await getSiteIdCached(user.token);
      if (!siteId) throw new Error("Site SharePoint ECOGEST introuvable");
      const res = await fetch(`https://graph.microsoft.com/v1.0/sites/${siteId}/lists?$top=10`, {
        headers: { Authorization: `Bearer ${user.token}` },
      });
      const data = await res.json();
      if (data.error) throw new Error(data.error.message);
      const lists = (data.value || []).map((l) => l.displayName);
      const ecoLists = lists.filter((l) => l.startsWith("ECOGEST_"));
      setTestResult({
        ok: true,
        msg: `✅ Connexion SharePoint réussie !`,
        detail: `${ecoLists.length} liste(s) ECOGEST trouvée(s) : ${ecoLists.join(", ")}`,
      });
    } catch (e) {
      setTestResult({ ok: false, msg: `❌ Erreur : ${e.message}` });
    }
    setTesting(false);
  };
  return (
    /* @__PURE__ */ <div>
      <PageHeader
        title="Configuration SharePoint"
        sub="Connexion à ecolesuperieurederabat.sharepoint.com/sites/ECOGEST"
      />
      <div className="eco-grid2" style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16 }}>
        <div>
          <Card style={{ padding: 24, marginBottom: 16 }}>
            <h4 style={{ margin: "0 0 16px", color: C.navy }}>Paramètres Azure AD</h4>
            <div
              style={{
                background: "#fff8f0",
                border: "1px solid #ffd9b8",
                borderRadius: 8,
                padding: "10px 14px",
                marginBottom: 16,
                fontSize: 12,
                color: "#a05020",
              }}
            >
              ✅ Connexion SharePoint activée — Administrateur approuvé → Azure Active Directory → Inscriptions
              d'applications
            </div>
            <Input
              label="URL SharePoint"
              value={cfg.siteUrl}
              onChange={(e) => setCfg((c) => ({ ...c, siteUrl: e.target.value }))}
            />
            <Input
              label="Tenant ID"
              value={cfg.tenantId}
              onChange={(e) => setCfg((c) => ({ ...c, tenantId: e.target.value }))}
              placeholder="xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
            />
            <Input
              label="Client ID"
              value={cfg.clientId}
              onChange={(e) => setCfg((c) => ({ ...c, clientId: e.target.value }))}
              placeholder="xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
            />
            <div
              style={{
                background: "#f0fff4",
                border: "1px solid #9ae6b4",
                borderRadius: 8,
                padding: "10px 14px",
                fontSize: 13,
                color: C.green,
                marginBottom: 8,
              }}
            >
              ✅ Connecté à SharePoint — Listes ECOGEST actives
            </div>
            {testResult && (
              /* @__PURE__ */ <div
                style={{
                  background: testResult.ok ? "#f0fff4" : "#fff5f5",
                  border: `1px solid ${testResult.ok ? "#9ae6b4" : "#feb2b2"}`,
                  borderRadius: 8,
                  padding: "10px 14px",
                  marginBottom: 14,
                  fontSize: 13,
                }}
              >
                <div style={{ fontWeight: 600, color: testResult.ok ? C.green : C.red }}>{testResult.msg}</div>
                {testResult.detail && (
                  /* @__PURE__ */ <div style={{ fontSize: 12, color: C.muted, marginTop: 4 }}>{testResult.detail}</div>
                )}
              </div>
            )}
            <div style={{ display: "flex", gap: 10 }}>
              <Btn
                onClick={() => {
                  setSaved(true);
                  toast("Configuration sauvegardée ✓");
                  setTimeout(() => setSaved(false), 2e3);
                }}
              >
                {saved ? "✓ Sauvegardé !" : "Sauvegarder"}
              </Btn>
              <Btn
                variant="light"
                onClick={async () => {
                  toast("Configuration des listes en cours...");
                  await setupSharePointLists(user?.token);
                  toast("✅ Listes SharePoint configurées !");
                }}
              >
                🔧 Configurer les listes
              </Btn>
              <Btn variant="navy" onClick={handleTest} disabled={testing}>
                {testing ? "⏳ Test en cours..." : "🔌 Tester la connexion"}
              </Btn>
            </div>
          </Card>
          <Card style={{ padding: 16 }}>
            <h4 style={{ margin: "0 0 12px", color: C.navy, fontSize: 14 }}>Statut de la connexion</h4>
            <div style={{ display: "flex", alignItems: "center", gap: 10, fontSize: 13 }}>
              <div style={{ width: 10, height: 10, borderRadius: "50%", background: user?.token ? C.green : C.red }} />
              <span style={{ color: user?.token ? C.green : C.red, fontWeight: 600 }}>
                {user?.token ? "Connecté via Microsoft 365" : "Non connecté"}
              </span>
            </div>
            {user && (
              /* @__PURE__ */ <div style={{ marginTop: 8, fontSize: 12, color: C.muted }}>
                <div>👤 {user.nom}</div>
                <div>📧 {user.email}</div>
                <div>🔑 Token Microsoft actif</div>
              </div>
            )}
          </Card>
        </div>
        <div>
          <Card style={{ padding: 24 }}>
            <h4 style={{ margin: "0 0 16px", color: C.navy }}>Listes SharePoint à créer</h4>
            <p style={{ fontSize: 12, color: C.muted, marginBottom: 14 }}>
              Créez ces listes dans votre site ECOGEST avec les colonnes correspondantes :
            </p>
            {[
              { nom: "ECOGEST_Etudiants", cols: "Nom, Prénom, Filière, Année, Email, Statut, Genre" },
              { nom: "ECOGEST_Professeurs", cols: "Nom, Prénom, CIN, Matière, Email, Type, Statut" },
              { nom: "ECOGEST_Notes", cols: "EtudiantId, Étudiant, Module, Note, Semestre, Coef, Date" },
              { nom: "ECOGEST_Presences", cols: "Étudiant, Module, Date, Heure, Statut" },
              { nom: "ECOGEST_Cours", cols: "Intitulé, Professeur, Filière, Semestre, Heures, Salle" },
              { nom: "ECOGEST_Annonces", cols: "Titre, Contenu, Auteur, Cible, Date" },
            ].map((l) => (
              /* @__PURE__ */ <div key={l.nom} style={{ padding: "10px 0", borderBottom: `1px solid ${C.light}` }}>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                  <div style={{ fontWeight: 600, fontSize: 13, color: C.navy }}>{l.nom}</div>
                  <Badge color={C.green} size={10}>
                    À créer
                  </Badge>
                </div>
                <div style={{ fontSize: 11, color: C.muted, marginTop: 3 }}>{l.cols}</div>
              </div>
            ))}
          </Card>
        </div>
      </div>
    </div>
  );
}
const getFilieres = (data) => {
  return (data.filieres || [])
    .filter((f) => f.statut === "Active")
    .map((f) => f.intitule)
    .filter(Boolean)
    .sort((a, b) => a.localeCompare(b, "fr", { numeric: true, sensitivity: "base" }));
};
const getTarifFiliere = (data, intitule) => {
  const f = (data.filieres || []).find((x) => x.intitule === intitule);
  return f ? f.frais : 0;
};
const DocButton = ({ icon, label, onClick, badge }) => (
  /* @__PURE__ */ <button
    onClick={onClick}
    title={label}
    style={{
      position: "relative",
      display: "flex",
      flexDirection: "column",
      alignItems: "center",
      justifyContent: "center",
      gap: 3,
      border: `1px solid ${C.border}`,
      background: C.light,
      borderRadius: 8,
      padding: "7px 4px",
      cursor: "pointer",
      fontFamily: "inherit",
      width: 66,
      transition: "background 0.15s, border-color 0.15s",
    }}
    onMouseEnter={(e) => {
      e.currentTarget.style.background = "#eef2fb";
      e.currentTarget.style.borderColor = C.navy;
    }}
    onMouseLeave={(e) => {
      e.currentTarget.style.background = C.light;
      e.currentTarget.style.borderColor = C.border;
    }}
  >
    {badge !== void 0 && badge !== null && (
      /* @__PURE__ */ <span
        style={{
          position: "absolute",
          top: -6,
          right: -6,
          background: C.orange,
          color: "white",
          fontSize: 9.5,
          fontWeight: 800,
          borderRadius: "50%",
          minWidth: 16,
          height: 16,
          padding: "0 3px",
          display: "flex",
          alignItems: "center",
          justifyContent: "center",
          boxShadow: "0 1px 3px rgba(0,0,0,0.25)",
        }}
      >
        {badge}
      </span>
    )}
    <span style={{ fontSize: 17 }}>{icon}</span>
    <span style={{ fontSize: 9.5, fontWeight: 700, color: C.navy, letterSpacing: "0.01em" }}>{label}</span>
  </button>
);
function Filieres({ data, setData, role, user, toast }) {
  const canEdit = ["administrateur"].includes(role);
  const [showModal, setShowModal] = useState(false);
  const [editItem, setEditItem] = useState(null);
  const [filterNiveau, setFilterNiveau] = useState("");
  const [search, setSearch] = useState("");
  const [showFiche, setShowFiche] = useState(null);
  const [showEtudiantsFiliere, setShowEtudiantsFiliere] = useState(null);
  const [form, setForm] = useState({
    code: "",
    intitule: "",
    niveau: "BAC+3",
    duree: 3,
    capacite: 30,
    frais: 18e3,
    statut: "Active",
    description: "",
    responsable: "",
    couleur: "#1a2f5e",
    typeSeances: "Soir",
  });
  const COULEURS = ["#1a2f5e", "#e56b2d", "#3b82f6", "#10b981", "#8b5cf6", "#ef4444", "#f59e0b", "#06b6d4", "#ec4899"];
  const filtered = (data.filieres || [])
    .filter((f) => {
      const matchNiveau = !filterNiveau || f.niveau === filterNiveau;
      const matchSearch = !search || `${f.code || ""} ${f.intitule || ""}`.toLowerCase().includes(search.toLowerCase());
      return matchNiveau && matchSearch;
    })
    .sort((a, b) => (a.code || "").localeCompare(b.code || "", "fr", { numeric: true, sensitivity: "base" }));
  const bac3 = (data.filieres || []).filter((f) => f.niveau === "BAC+3");
  const bac5 = (data.filieres || []).filter((f) => f.niveau === "BAC+5");
  const openAdd = () => {
    setEditItem(null);
    setForm({
      code: "",
      intitule: "",
      niveau: "BAC+3",
      duree: 3,
      capacite: 30,
      frais: 18e3,
      statut: "Active",
      description: "",
      responsable: "",
      couleur: "#1a2f5e",
      typeSeances: "Soir",
    });
    setShowModal(true);
  };
  const openEdit = (f) => {
    setEditItem(f);
    setForm({ ...f });
    setShowModal(true);
  };
  const spFields = (f) => ({
    Title: f.intitule || f.code || "",
    Code: f.code || "",
    Intitule: f.intitule || "",
    Niveau: f.niveau || "BAC+3",
    Duree: String(f.duree || 3),
    Capacite: String(f.capacite || 30),
    Frais: String(f.frais || 0),
    Statut: f.statut || "Active",
    Description: f.description || "",
    Responsable: f.responsable || "",
    Couleur: f.couleur || "#1a2f5e",
    TypeSeances: f.typeSeances || "Soir",
  });
  const handleSave = () => {
    if (editItem) {
      setData((d) => ({
        ...d,
        filieres: d.filieres.map((f) =>
          f.id === editItem.id ? { ...form, id: editItem.id, _spId: editItem._spId } : f
        ),
      }));
      if (user?.token && editItem._spId) saveToSP("ECOGEST_Filieres", spFields(form), user.token, editItem._spId);
      else if (user?.token && !editItem._spId) console.warn("Modification filière: _spId manquant — rechargez la page");
      toast("Groupe mis à jour dans SharePoint ✓");
      setShowFiche((f) => (f && f.id === editItem.id ? { ...form, id: editItem.id, _spId: editItem._spId } : f));
    } else {
      const newF = { ...form, id: Date.now() };
      if (!user?.token) {
        toast("⚠️ Non connecté à SharePoint — groupe non enregistré");
      } else {
        saveToSP("ECOGEST_Filieres", spFields(form), user.token).then((result) => {
          if (!result) {
            toast("⚠️ Échec de l'enregistrement dans SharePoint — groupe non ajouté");
            return;
          }
          setData((d) => ({ ...d, filieres: [...(d.filieres || []), { ...newF, _spId: result.id }] }));
          toast("Groupe ajouté dans SharePoint ✓");
        });
      }
    }
    setShowModal(false);
  };
  const handleToggle = (id) => {
    const fil = (data.filieres || []).find((f) => f.id === id);
    if (!fil) return;
    const newStatut = fil.statut === "Active" ? "Inactive" : "Active";
    setData((d) => ({ ...d, filieres: d.filieres.map((f) => (f.id === id ? { ...f, statut: newStatut } : f)) }));
    if (user?.token && fil._spId) saveToSP("ECOGEST_Filieres", { Statut: newStatut }, user.token, fil._spId);
    toast("Statut mis à jour dans SharePoint ✓");
    setShowFiche((f) => (f && f.id === id ? { ...f, statut: newStatut } : f));
  };
  const handleDelete = (id) => {
    if (!window.confirm("Supprimer ce groupe ?")) return;
    const fil = (data.filieres || []).find((f) => f.id === id);
    setData((d) => ({ ...d, filieres: d.filieres.filter((f) => f.id !== id) }));
    if (user?.token && fil?._spId) deleteFromSP("ECOGEST_Filieres", fil._spId, user.token);
    toast("Groupe supprimé de SharePoint ✓");
    setShowFiche((f) => (f && f.id === id ? null : f));
  };
  const nbEtudiants = (intitule) => (data.etudiants || []).filter((e) => e.filiere === intitule).length;
  const nbVisiteurs = (intitule) => (data.visiteurs || []).filter((v) => v.filiereS === intitule).length;
  return (
    /* @__PURE__ */ <div>
      <PageHeader
        title="Gestion des Groupes"
        sub={`${(data.filieres || []).length} groupe(s) — liste SharePoint "${SP.lists.filieres}"`}
        action={canEdit && /* @__PURE__ */ <Btn onClick={openAdd}>+ Nouveau groupe</Btn>}
      />
      <div style={{ display: "flex", gap: 12, marginBottom: 20 }}>
        <StatCard label="Total groupes" value={(data.filieres || []).length} color={C.navy} icon="📚" />
        <StatCard label="BAC+3" value={bac3.length} color={C.teal} icon="🎓" />
        <StatCard label="BAC+5" value={bac5.length} color={C.purple} icon="🏆" />
        <StatCard
          label="Groupes actifs"
          value={(data.filieres || []).filter((f) => f.statut === "Active").length}
          color={C.green}
          icon="✅"
        />
      </div>
      <div style={{ display: "flex", gap: 10, marginBottom: 14, flexWrap: "wrap" }}>
        <input
          placeholder="🔍 Rechercher par code ou intitulé..."
          value={search}
          onChange={(e) => setSearch(e.target.value)}
          style={{
            flex: 1,
            minWidth: 200,
            padding: "9px 14px",
            border: `1px solid ${C.border}`,
            borderRadius: 8,
            fontSize: 13,
            fontFamily: "inherit",
          }}
        />
        {["", "BAC+3", "BAC+5"].map((n) => (
          /* @__PURE__ */ <button
            key={n}
            onClick={() => setFilterNiveau(n)}
            style={{
              padding: "8px 14px",
              borderRadius: 20,
              border: `1px solid ${filterNiveau === n ? C.navy : C.border}`,
              background: filterNiveau === n ? C.navy : C.white,
              color: filterNiveau === n ? C.white : "#555",
              fontSize: 12,
              cursor: "pointer",
              fontFamily: "inherit",
              fontWeight: 600,
              boxShadow: filterNiveau === n ? "0 3px 10px rgba(26,47,94,0.28)" : "none",
              transition: "all 0.13s",
            }}
          >
            {n || "Tous"}
          </button>
        ))}
      </div>
      <Card style={{ overflow: "hidden" }}>
        <DataTable
          cols={[
            {
              key: "code",
              label: "Code",
              render: (v, row) => (
                /* @__PURE__ */ <span
                  style={{
                    background: colorForAnnee(row.intitule),
                    color: "#ffffff",
                    fontWeight: 800,
                    fontSize: 11.5,
                    padding: "4px 11px",
                    borderRadius: 20,
                    boxShadow: `0 2px 6px ${colorForAnnee(row.intitule)}66`,
                    display: "inline-block",
                    letterSpacing: "0.02em",
                    whiteSpace: "nowrap",
                  }}
                >
                  {v || "—"}
                </span>
              ),
            },
            { key: "intitule", label: "Intitulé", render: (v, row) => v || row.code || row.Title || "—" },
            {
              key: "niveau",
              label: "Niveau",
              render: (v) => (
                /* @__PURE__ */ <Badge color={v === "BAC+5" ? C.purple : C.teal}>
                  {v === "BAC+5" ? "🏆" : "🎓"} {v}
                </Badge>
              ),
            },
            {
              key: "statut",
              label: "Statut",
              render: (v) => /* @__PURE__ */ <Badge color={v === "Active" ? C.green : C.muted}>{v}</Badge>,
            },
            {
              key: "_etudiants",
              label: "Étudiants",
              render: (_, row) => (
                /* @__PURE__ */ <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                  <span style={{ fontWeight: 800, color: C.navy }}>🎓 {nbEtudiants(row.intitule)}</span>
                  {nbEtudiants(row.intitule) > 0 && (
                    /* @__PURE__ */ <button
                      onClick={() => setShowEtudiantsFiliere(row)}
                      style={{
                        border: "none",
                        background: `linear-gradient(135deg, ${C.navy} 0%, ${C.teal} 100%)`,
                        color: "white",
                        fontSize: 11,
                        fontWeight: 700,
                        padding: "6px 14px",
                        borderRadius: 20,
                        cursor: "pointer",
                        fontFamily: "inherit",
                        whiteSpace: "nowrap",
                        boxShadow: "0 3px 10px rgba(26,47,94,0.4)",
                        transition: "transform 0.15s, box-shadow 0.15s, filter 0.15s",
                        display: "inline-flex",
                        alignItems: "center",
                        gap: 5,
                      }}
                      onMouseEnter={(e) => {
                        e.currentTarget.style.transform = "translateY(-2px)";
                        e.currentTarget.style.boxShadow = "0 5px 14px rgba(26,47,94,0.5)";
                        e.currentTarget.style.filter = "brightness(1.06)";
                      }}
                      onMouseLeave={(e) => {
                        e.currentTarget.style.transform = "translateY(0)";
                        e.currentTarget.style.boxShadow = "0 3px 10px rgba(26,47,94,0.4)";
                        e.currentTarget.style.filter = "brightness(1)";
                      }}
                    >
                      👁️ Voir la liste
                    </button>
                  )}
                </div>
              ),
            },
            {
              key: "_visiteurs",
              label: "Visiteurs",
              render: (_, row) => (
                /* @__PURE__ */ <span style={{ fontWeight: 800, color: C.orange }}>👋 {nbVisiteurs(row.intitule)}</span>
              ),
            },
            ...(canEdit
              ? [
                  {
                    key: "_actions",
                    label: "Actions",
                    render: (_, row) => (
                      /* @__PURE__ */ <div style={{ display: "flex", gap: 6 }}>
                        <button
                          title="Modifier"
                          onClick={() => openEdit(row)}
                          style={{
                            border: "none",
                            borderRadius: 8,
                            cursor: "pointer",
                            width: 32,
                            height: 32,
                            fontSize: 14,
                            background: C.navy,
                            color: "white",
                          }}
                        >
                          ✏️
                        </button>
                        <button
                          title={row.statut === "Active" ? "Désactiver" : "Activer"}
                          onClick={() => handleToggle(row.id)}
                          style={{
                            border: "none",
                            borderRadius: 8,
                            cursor: "pointer",
                            width: 32,
                            height: 32,
                            fontSize: 14,
                            background: row.statut === "Active" ? C.light : C.orange,
                            color: row.statut === "Active" ? "#555" : "white",
                          }}
                        >
                          {row.statut === "Active" ? "⏸" : "▶"}
                        </button>
                        <button
                          title="Supprimer"
                          onClick={() => handleDelete(row.id)}
                          style={{
                            border: "none",
                            borderRadius: 8,
                            cursor: "pointer",
                            width: 32,
                            height: 32,
                            fontSize: 14,
                            background: C.red,
                            color: "white",
                          }}
                        >
                          🗑
                        </button>
                      </div>
                    ),
                  },
                ]
              : []),
          ]}
          rows={filtered}
          emptyMsg="Aucun groupe ne correspond à votre recherche"
          compact={true}
          rowStyle={(row, i) => (i % 2 === 0 ? C.navyLight : "#e0f4f2")}
        />
      </Card>
      {showModal && (
        /* @__PURE__ */ <Modal
          title={editItem ? "Modifier le groupe" : "Nouveau groupe → SharePoint"}
          onClose={() => setShowModal(false)}
          width={600}
        >
          <div className="eco-grid2" style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "0 14px" }}>
            <Input
              label="Code *"
              value={form.code}
              onChange={(e) => setForm((f) => ({ ...f, code: e.target.value }))}
              placeholder="ex: MGE, SI-GL..."
            />
            <Sel
              label="Niveau *"
              value={form.niveau}
              onChange={(e) => setForm((f) => ({ ...f, niveau: e.target.value }))}
            >
              <option>BAC+3</option>
              <option>BAC+5</option>
            </Sel>
          </div>
          <Input
            label="Intitulé complet *"
            value={form.intitule}
            onChange={(e) => setForm((f) => ({ ...f, intitule: e.target.value }))}
            placeholder="BAC+3 en ..."
          />
          <div style={{ marginBottom: 13 }}>
            <label style={{ fontSize: 12, fontWeight: 600, color: "#555", display: "block", marginBottom: 4 }}>
              Description
            </label>
            <textarea
              value={form.description}
              onChange={(e) => setForm((f) => ({ ...f, description: e.target.value }))}
              rows={2}
              style={{
                width: "100%",
                padding: "8px 11px",
                border: `1px solid ${C.border}`,
                borderRadius: 8,
                fontSize: 13,
                fontFamily: "inherit",
                resize: "vertical",
                boxSizing: "border-box",
              }}
            />
          </div>
          <Sel label="Statut" value={form.statut} onChange={(e) => setForm((f) => ({ ...f, statut: e.target.value }))}>
            <option>Active</option>
            <option>Inactive</option>
          </Sel>
          <Sel
            label="Type de séances (Emploi du temps)"
            value={form.typeSeances}
            onChange={(e) => setForm((f) => ({ ...f, typeSeances: e.target.value }))}
          >
            <option value="Soir">🌙 Soir uniquement (19h00 - 21h00)</option>
            <option value="Jour">☀️ Journée (tous les créneaux sauf 19h00 - 21h00)</option>
          </Sel>
          <Input
            label="Responsable de filière"
            value={form.responsable}
            onChange={(e) => setForm((f) => ({ ...f, responsable: e.target.value }))}
            placeholder="Nom du responsable"
          />
          <div style={{ marginBottom: 14 }}>
            <label style={{ fontSize: 12, fontWeight: 600, color: "#555", display: "block", marginBottom: 6 }}>
              Couleur d'identification
            </label>
            <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
              {COULEURS.map((c) => (
                /* @__PURE__ */ <button
                  key={c}
                  onClick={() => setForm((f) => ({ ...f, couleur: c }))}
                  style={{
                    width: 28,
                    height: 28,
                    borderRadius: "50%",
                    background: c,
                    border: `3px solid ${form.couleur === c ? "#333" : "transparent"}`,
                    cursor: "pointer",
                    transition: "transform 0.1s",
                    transform: form.couleur === c ? "scale(1.2)" : "scale(1)",
                  }}
                />
              ))}
            </div>
          </div>
          <div
            style={{
              background: "#f0f7ff",
              border: "1px solid #bee3f8",
              borderRadius: 8,
              padding: "8px 14px",
              fontSize: 12,
              color: "#2b6cb0",
              marginBottom: 14,
            }}
          >
            📋 Données enregistrées dans la liste <strong>"{SP.lists.filieres}"</strong> sur SharePoint ECOGEST
          </div>
          <div style={{ display: "flex", gap: 10, justifyContent: "flex-end" }}>
            <Btn variant="light" onClick={() => setShowModal(false)}>
              Annuler
            </Btn>
            <Btn onClick={handleSave} disabled={!form.code || !form.intitule}>
              {editItem ? "Mettre à jour → SharePoint" : "Ajouter → SharePoint"}
            </Btn>
          </div>
        </Modal>
      )}
      {showFiche && (
        /* @__PURE__ */ <Modal title="Fiche du groupe" onClose={() => setShowFiche(null)} width={520}>
          <div style={{ margin: "-28px -28px 20px", borderRadius: "16px 16px 0 0", overflow: "hidden" }}>
            <div
              style={{
                background: `linear-gradient(135deg, ${showFiche.couleur} 0%, ${showFiche.couleur}c8 100%)`,
                padding: "22px 26px",
                display: "flex",
                justifyContent: "space-between",
                alignItems: "flex-start",
              }}
            >
              <div>
                <div
                  style={{
                    display: "inline-flex",
                    alignItems: "center",
                    gap: 5,
                    fontSize: 10.5,
                    fontWeight: 700,
                    color: "rgba(255,255,255,0.9)",
                    textTransform: "uppercase",
                    letterSpacing: "0.08em",
                    marginBottom: 8,
                    background: "rgba(255,255,255,0.18)",
                    padding: "3px 10px",
                    borderRadius: 20,
                  }}
                >
                  {showFiche.niveau === "BAC+5" ? "🏆" : "🎓"} {showFiche.code ? `${showFiche.code} · ${showFiche.niveau || ""}` : showFiche.niveau || ""}
                </div>
                <div
                  style={{ fontSize: 17, fontWeight: 800, color: "white", lineHeight: 1.35, letterSpacing: "-0.01em" }}
                >
                  {showFiche.intitule || showFiche.code || "—"}
                </div>
              </div>
              <Badge color="rgba(255,255,255,0.95)" size={11}>
                {showFiche.statut}
              </Badge>
            </div>
          </div>
          {showFiche.description && (
            /* @__PURE__ */ <p style={{ fontSize: 13, color: "#555", margin: "0 0 16px", lineHeight: 1.6 }}>
              {showFiche.description}
            </p>
          )}
          <div style={{ display: "flex", gap: 10, marginBottom: 16 }}>
            <div
              onClick={() => nbEtudiants(showFiche.intitule) > 0 && setShowEtudiantsFiliere(showFiche)}
              style={{
                flex: 1,
                background: C.navyLight,
                borderRadius: 10,
                padding: "12px 10px",
                textAlign: "center",
                cursor: nbEtudiants(showFiche.intitule) > 0 ? "pointer" : "default",
              }}
              title={nbEtudiants(showFiche.intitule) > 0 ? "Voir la liste des étudiants inscrits" : ""}
            >
              <div style={{ fontSize: 10.5, color: C.muted, fontWeight: 600, marginBottom: 3 }}>
                🎓 Étudiants inscrits
              </div>
              <div
                style={{
                  fontSize: 20,
                  fontWeight: 800,
                  color: C.navy,
                  textDecoration: nbEtudiants(showFiche.intitule) > 0 ? "underline" : "none",
                  textUnderlineOffset: 3,
                }}
              >
                {nbEtudiants(showFiche.intitule)}
              </div>
            </div>
            <div
              style={{
                flex: 1,
                background: C.orangeLight,
                borderRadius: 10,
                padding: "12px 10px",
                textAlign: "center",
              }}
            >
              <div style={{ fontSize: 10.5, color: C.muted, fontWeight: 600, marginBottom: 3 }}>
                👋 Visiteurs intéressés
              </div>
              <div style={{ fontSize: 20, fontWeight: 800, color: C.orange }}>{nbVisiteurs(showFiche.intitule)}</div>
            </div>
          </div>
          {showFiche.responsable && (
            /* @__PURE__ */ <div style={{ fontSize: 13, color: "#555", marginBottom: 16 }}>
              <b>Responsable :</b> {showFiche.responsable}
            </div>
          )}
          {canEdit && (
            /* @__PURE__ */ <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
              <Btn
                small={true}
                variant="navy"
                onClick={() => {
                  setShowFiche(null);
                  openEdit(showFiche);
                }}
                style={{ flex: 1 }}
              >
                ✏️ Modifier
              </Btn>
              <Btn
                small={true}
                variant={showFiche.statut === "Active" ? "light" : "primary"}
                onClick={() => handleToggle(showFiche.id)}
                style={{ flex: 1 }}
              >
                {showFiche.statut === "Active" ? "⏸ Désactiver" : "▶ Activer"}
              </Btn>
              <Btn
                small={true}
                variant="danger"
                onClick={() => {
                  setShowFiche(null);
                  handleDelete(showFiche.id);
                }}
              >
                🗑
              </Btn>
            </div>
          )}
        </Modal>
      )}
      {showEtudiantsFiliere && (
        /* @__PURE__ */ <Modal
          title={`Étudiants — ${showEtudiantsFiliere.intitule}`}
          onClose={() => setShowEtudiantsFiliere(null)}
          width={800}
        >
          {(data.etudiants || []).filter((e) => e.filiere === showEtudiantsFiliere.intitule).length === 0 ? (
            /* @__PURE__ */ <div style={{ textAlign: "center", padding: 32, color: C.muted, fontSize: 13 }}>
              Aucun étudiant inscrit dans ce groupe
            </div>
          ) : (
            (data.etudiants || [])
              .filter((e) => e.filiere === showEtudiantsFiliere.intitule)
              .map((et) => (
                /* @__PURE__ */ <div
                  key={et.id}
                  style={{
                    display: "flex",
                    alignItems: "center",
                    gap: 20,
                    padding: "14px 6px",
                    borderBottom: `1px solid ${C.light}`,
                  }}
                >
                  <div style={{ minWidth: 0, flexShrink: 0, width: 190 }}>
                    <div style={{ fontWeight: 700, fontSize: 14, color: C.navy }}>
                      {et.prenom} {et.nom}
                    </div>
                    <div
                      style={{
                        fontSize: 12,
                        color: C.muted,
                        marginTop: 3,
                        display: "flex",
                        alignItems: "center",
                        gap: 6,
                        flexWrap: "wrap",
                      }}
                    >
                      <span>{et.numInscription || "—"}</span>
                      <Badge color={et.statut === "Actif" ? C.green : C.muted}>{et.statut}</Badge>
                    </div>
                  </div>
                  <div
                    style={{
                      flex: 1,
                      minWidth: 0,
                      fontSize: 12,
                      color: "#555",
                      display: "flex",
                      flexDirection: "column",
                      gap: 4,
                    }}
                  >
                    <div
                      style={{
                        display: "flex",
                        alignItems: "center",
                        gap: 6,
                        overflow: "hidden",
                        textOverflow: "ellipsis",
                        whiteSpace: "nowrap",
                      }}
                    >
                      <span style={{ color: C.muted }}>✉️</span> {et.email || "—"}
                    </div>
                    <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
                      <span style={{ color: C.muted }}>📞</span> {et.telephone || "—"}
                    </div>
                  </div>
                  <div style={{ display: "flex", gap: 8, flexShrink: 0 }}>
                    <DocButton
                      icon="📄"
                      label="Inscription"
                      onClick={() => genererAttestation(et, showEtudiantsFiliere.intitule, "inscription")}
                    />
                    <DocButton
                      icon="📜"
                      label="Scolarité"
                      onClick={() => genererAttestation(et, showEtudiantsFiliere.intitule, "scolarite")}
                    />
                    <DocButton
                      icon="🏆"
                      label="Réussite"
                      onClick={() => genererAttestation(et, showEtudiantsFiliere.intitule, "reussite")}
                    />
                  </div>
                </div>
              ))
          )}
          <div
            style={{
              background: "#f0f7ff",
              border: "1px solid #bee3f8",
              borderRadius: 8,
              padding: "8px 14px",
              fontSize: 11.5,
              color: "#2b6cb0",
              marginTop: 14,
            }}
          >
            📄 Les documents sont générés au format PDF et téléchargés directement sur votre appareil.
          </div>
        </Modal>
      )}
    </div>
  );
}
function FicheVisiteur({ visiteur, onClose, onConvertir, onSave, onDelete, canEdit, user, filieresList }) {
  const [editMode, setEditMode] = React.useState(false);
  const [form, setForm] = React.useState(() => ({ ...visiteur }));
  React.useEffect(() => {
    setForm({ ...visiteur });
    setEditMode(false);
  }, [visiteur.id]);
  const setF = (k, v) => setForm((f) => ({ ...f, [k]: v }));
  const handleSaveEdit = () => {
    onSave({ ...visiteur, ...form });
    setEditMode(false);
  };
  const handleCancelEdit = () => {
    setForm({ ...visiteur });
    setEditMode(false);
  };
  return (
    /* @__PURE__ */ <Modal title={`Fiche Visiteur — ${visiteur.prenom} ${visiteur.nom}`} onClose={onClose} width={680}>
      <div
        style={{
          display: "flex",
          justifyContent: "space-between",
          alignItems: "center",
          marginBottom: 20,
          padding: "10px 14px",
          background:
            visiteur.statut === "Inscrit" ? "#f0fff4" : visiteur.statut === "Non intéressé" ? "#fff5f5" : "#f0f7ff",
          borderRadius: 8,
          border: `1px solid ${visiteur.statut === "Inscrit" ? "#9ae6b4" : visiteur.statut === "Non intéressé" ? "#feb2b2" : "#bee3f8"}`,
        }}
      >
        <div
          style={{
            fontSize: 13,
            fontWeight: 600,
            color: visiteur.statut === "Inscrit" ? C.green : visiteur.statut === "Non intéressé" ? C.red : C.navy,
          }}
        >
          Statut : {visiteur.statut}
        </div>
        {canEdit && editMode && (
          /* @__PURE__ */ <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
            <Btn small={true} variant="light" onClick={handleCancelEdit}>
              ✕ Annuler
            </Btn>
            <Btn
              small={true}
              variant="primary"
              onClick={handleSaveEdit}
              disabled={!form.nom || !form.prenom || !form.cin || !/^\d{10}$/.test(form.telephone || "")}
            >
              💾 Enregistrer
            </Btn>
          </div>
        )}
        {canEdit && !editMode && (
          /* @__PURE__ */ <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
            <Btn small={true} variant="navy" onClick={() => setEditMode(true)}>
              ✏️ Modifier
            </Btn>
            {visiteur.statut === "Visiteur" && (
              /* @__PURE__ */ <React.Fragment>
                <Btn small={true} variant="primary" onClick={() => onConvertir(visiteur.id, "Inscrit")}>
                  ✓ Convertir en Étudiant
                </Btn>
                <Btn small={true} variant="danger" onClick={() => onConvertir(visiteur.id, "Non intéressé")}>
                  ✗ Non intéressé
                </Btn>
              </React.Fragment>
            )}
            <Btn small={true} variant="danger" onClick={() => onDelete(visiteur)}>
              🗑 Supprimer
            </Btn>
          </div>
        )}
        {!canEdit && visiteur.statut === "Inscrit" && (
          /* @__PURE__ */ <Badge color={C.green}>✓ Converti en Étudiant</Badge>
        )}
      </div>
      <div style={{ marginBottom: 16 }}>
        <div
          style={{
            fontSize: 11,
            fontWeight: 700,
            color: C.orange,
            textTransform: "uppercase",
            letterSpacing: "0.08em",
            marginBottom: 10,
          }}
        >
          Informations personnelles
        </div>
        {!editMode ? (
          /* @__PURE__ */ <div
            className="eco-grid3"
            style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: "6px 14px" }}
          >
            {[
              ["Nom", visiteur.nom],
              ["Prénom", visiteur.prenom],
              ["CIN", visiteur.cin],
              ["Date naissance", formatDateFR(visiteur.dateNaissance)],
              ["Lieu naissance", visiteur.lieuNaissance],
              ["Genre", visiteur.genre === "F" ? "Féminin" : "Masculin"],
              ["Téléphone", visiteur.telephone],
              ["Email", visiteur.email],
              ["Ville", visiteur.ville],
              ["Quartier", visiteur.quartier],
              ["Adresse", visiteur.adresse],
              ["Source", visiteur.source],
              ["Reçu par", visiteur.recuPar],
            ].map(([lbl, val]) => (
              /* @__PURE__ */ <div key={lbl} style={{ background: C.light, borderRadius: 6, padding: "7px 10px" }}>
                <div style={{ fontSize: 10, color: C.muted, marginBottom: 2 }}>{lbl}</div>
                <div style={{ fontSize: 12, fontWeight: 600, color: C.navy }}>{val || "—"}</div>
              </div>
            ))}
          </div>
        ) : (
          /* @__PURE__ */ <React.Fragment>
            <div className="eco-grid3" style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: "0 14px" }}>
              <Input label="Nom *" value={form.nom || ""} onChange={(e) => setF("nom", e.target.value)} />
              <Input label="Prénom *" value={form.prenom || ""} onChange={(e) => setF("prenom", e.target.value)} />
              <Input label="CIN *" value={form.cin || ""} onChange={(e) => setF("cin", e.target.value)} />
            </div>
            <div className="eco-grid2" style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "0 14px" }}>
              <Input
                label="Date de naissance"
                type="date"
                value={form.dateNaissance || ""}
                onChange={(e) => setF("dateNaissance", e.target.value)}
              />
              <Input
                label="Lieu de naissance"
                value={form.lieuNaissance || ""}
                onChange={(e) => setF("lieuNaissance", e.target.value)}
              />
            </div>
            <div className="eco-grid3" style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: "0 14px" }}>
              <Sel label="Genre" value={form.genre || "M"} onChange={(e) => setF("genre", e.target.value)}>
                <option value="M">Masculin</option>
                <option value="F">Féminin</option>
              </Sel>
              <Input
                label="Téléphone *"
                value={form.telephone || ""}
                onChange={(e) => setF("telephone", e.target.value)}
                placeholder="06XXXXXXXX"
                style={form.telephone && !/^\d{10}$/.test(form.telephone) ? { borderColor: C.red } : {}}
              />
              <Input
                label="Email"
                type="email"
                value={form.email || ""}
                onChange={(e) => setF("email", e.target.value)}
              />
            </div>
            <div className="eco-grid3" style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: "0 14px" }}>
              <Input label="Adresse" value={form.adresse || ""} onChange={(e) => setF("adresse", e.target.value)} />
              <Input label="Ville" value={form.ville || ""} onChange={(e) => setF("ville", e.target.value)} />
              <Input label="Quartier" value={form.quartier || ""} onChange={(e) => setF("quartier", e.target.value)} />
            </div>
            <div className="eco-grid2" style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "0 14px" }}>
              <Sel
                label="Source"
                value={form.source || "Visite directe"}
                onChange={(e) => setF("source", e.target.value)}
              >
                <option>Visite directe</option>
                <option>Téléphone</option>
                <option>Email</option>
                <option>Réseaux sociaux</option>
                <option>Recommandation</option>
                <option>Salon</option>
              </Sel>
              <Sel label="Reçu par" value={form.recuPar || ""} onChange={(e) => setF("recuPar", e.target.value)}>
                <option value="">-- Sélectionner --</option>
                <option>ADILA BAHALLA</option>
                <option>KAWTAR BOUABID</option>
                <option>FATIMA ZAHRA ZNIBER</option>
              </Sel>
            </div>
          </React.Fragment>
        )}
      </div>
      <div style={{ marginBottom: 16 }}>
        <div
          style={{
            fontSize: 11,
            fontWeight: 700,
            color: C.orange,
            textTransform: "uppercase",
            letterSpacing: "0.08em",
            marginBottom: 10,
          }}
        >
          Baccalauréat & Formation
        </div>
        {!editMode ? (
          /* @__PURE__ */ <div
            className="eco-grid3"
            style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: "6px 14px" }}
          >
            {[
              ["Année Bac", visiteur.bacAnnee],
              ["Mention", visiteur.bacMention],
              ["Série", visiteur.bacSerie],
              ["Établissement", visiteur.etablissement],
              ["Filière souhaitée", visiteur.filiereS],
              ["Date visite", formatDateFR(visiteur.dateVisite)],
            ].map(([lbl, val]) => (
              /* @__PURE__ */ <div key={lbl} style={{ background: C.light, borderRadius: 6, padding: "7px 10px" }}>
                <div style={{ fontSize: 10, color: C.muted, marginBottom: 2 }}>{lbl}</div>
                <div style={{ fontSize: 12, fontWeight: 600, color: C.navy }}>{val || "—"}</div>
              </div>
            ))}
          </div>
        ) : (
          /* @__PURE__ */ <React.Fragment>
            <div className="eco-grid3" style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: "0 14px" }}>
              <Input
                label="Année du Bac"
                value={form.bacAnnee || ""}
                onChange={(e) => setF("bacAnnee", e.target.value)}
                placeholder="2024"
              />
              <Sel
                label="Mention"
                value={form.bacMention || "Passable"}
                onChange={(e) => setF("bacMention", e.target.value)}
              >
                <option>Très bien</option>
                <option>Bien</option>
                <option>Assez bien</option>
                <option>Passable</option>
              </Sel>
              <Input
                label="Série du Bac"
                value={form.bacSerie || ""}
                onChange={(e) => setF("bacSerie", e.target.value)}
                placeholder="Sciences Éco..."
              />
            </div>
            <div className="eco-grid3" style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: "0 14px" }}>
              <Input
                label="Établissement d'origine"
                value={form.etablissement || ""}
                onChange={(e) => setF("etablissement", e.target.value)}
              />
              <Sel
                label="Filière souhaitée"
                value={form.filiereS || ""}
                onChange={(e) => setF("filiereS", e.target.value)}
              >
                <option value="">-- Sélectionner une filière --</option>
                {(filieresList || []).map((f) => (
                  /* @__PURE__ */ <option key={f} value={f}>
                    {f}
                  </option>
                ))}
              </Sel>
              <Input
                label="Date de visite"
                type="date"
                value={form.dateVisite || ""}
                onChange={(e) => setF("dateVisite", e.target.value)}
              />
            </div>
          </React.Fragment>
        )}
      </div>
    </Modal>
  );
}
function Visiteurs({ data, setData, role, user, toast }) {
  const canEdit = ["administrateur", "directrice"].includes(role);
  const [showModal, setShowModal] = useState(false);
  const [showFiche, setShowFiche] = useState(null);
  const [editVisiteur, setEditVisiteur] = useState(null);
  const [search, setSearch] = useState("");
  const [filterStatut, setFilterStatut] = useState("");
  const [form, setForm] = useState({
    nom: "",
    prenom: "",
    cin: "",
    dateNaissance: "",
    lieuNaissance: "",
    genre: "M",
    telephone: "",
    email: "",
    adresse: "",
    ville: "",
    quartier: "",
    bacAnnee: "",
    bacMention: "Passable",
    bacSerie: "",
    etablissement: "",
    filiereS: "",
    dateVisite: /* @__PURE__ */ new Date().toISOString().split("T")[0],
    source: "Visite directe",
    statut: "Visiteur",
    observations: "",
    recuPar: "",
    anneeUniversitaire: data.configuration?.anneeUniversitaireActive || ANNEES_UNIVERSITAIRES[0],
    documents: [],
    tarif: [],
  });
  const [filieresList, setFilieresList] = useState(getFilieres(data));
  useEffect(() => {
    const fl = getFilieres(data);
    setFilieresList(fl);
  }, [data.filieres]);
  const genTarif = (filiere) => {
    const total = getTarifFiliere(data, filiere);
    const tranche = Math.round(total / 3);
    return [
      { tranche: "1ère tranche", montant: tranche, echeance: "2026-09-15", paye: false },
      { tranche: "2ème tranche", montant: tranche, echeance: "2027-01-15", paye: false },
      { tranche: "3ème tranche", montant: total - tranche * 2, echeance: "2027-04-15", paye: false },
    ];
  };
  const filtered = (data.visiteurs || []).filter((v) => {
    const matchSearch = `${v.nom} ${v.prenom} ${v.cin} ${v.telephone}`.toLowerCase().includes(search.toLowerCase());
    const matchStatut = !filterStatut || v.statut === filterStatut;
    return matchSearch && matchStatut;
  });
  const filteredSortedVisiteurs = [...filtered].sort((a, b) => {
    const codeA = filiereCode(a.filiereS, data.filieres) || a.filiereS || "";
    const codeB = filiereCode(b.filiereS, data.filieres) || b.filiereS || "";
    return (
      codeA.localeCompare(codeB, "fr", { numeric: true, sensitivity: "base" }) ||
      (a.nom || "").localeCompare(b.nom || "", "fr", { sensitivity: "base" })
    );
  });
  const bandeFiliereVisParId = {};
  const premiereLigneFiliereVisId = {};
  const derniereLigneFiliereVisId = {};
  {
    let bande = 0,
      filierePrecedente = null;
    filteredSortedVisiteurs.forEach((e, idx) => {
      const code = filiereCode(e.filiereS, data.filieres) || e.filiereS || "";
      if (code !== filierePrecedente) {
        bande = 1 - bande;
        filierePrecedente = code;
        premiereLigneFiliereVisId[e.id] = true;
        if (idx > 0) derniereLigneFiliereVisId[filteredSortedVisiteurs[idx - 1].id] = true;
      }
      bandeFiliereVisParId[e.id] = bande;
    });
    if (filteredSortedVisiteurs.length > 0)
      derniereLigneFiliereVisId[filteredSortedVisiteurs[filteredSortedVisiteurs.length - 1].id] = true;
  }
  const handleSave = () => {
    const anneeActive = data.configuration?.anneeUniversitaireActive || ANNEES_UNIVERSITAIRES[0];
    const newV = {
      ...form,
      id: Date.now(),
      tarif: [],
      anneeUniversitaire: editVisiteur?.anneeUniversitaire || anneeActive,
    };
    const spF = {
      Title: (form.nom + " " + form.prenom).toUpperCase(),
      Nom: (form.nom || "").toUpperCase(),
      Prenom: (form.prenom || "").toUpperCase(),
      CIN: (form.cin || "").toUpperCase(),
      DateNaissance: form.dateNaissance || "",
      LieuNaissance: form.lieuNaissance || "",
      Genre: form.genre || "M",
      Telephone: form.telephone || "",
      Email: form.email || "",
      Adresse: form.adresse || "",
      Ville: form.ville || "",
      Quartier: form.quartier || "",
      BacAnnee: form.bacAnnee || "",
      BacMention: form.bacMention || "",
      BacSerie: form.bacSerie || "",
      Etablissement: form.etablissement || "",
      FiliereS: form.filiereS || "",
      DateVisite: form.dateVisite || "",
      Source: form.source || "Visite directe",
      Statut: form.statut || "Visiteur",
      Observations: form.observations || "",
      AnneeUniversitaire: newV.anneeUniversitaire,
      RecuPar: form.recuPar || "",
    };
    if (editVisiteur) {
      setData((d) => ({
        ...d,
        visiteurs: (d.visiteurs || []).map((v) =>
          v.id === editVisiteur.id ? { ...newV, id: editVisiteur.id, _spId: editVisiteur._spId } : v
        ),
      }));
      if (user?.token && editVisiteur._spId) saveToSP("ECOGEST_Visiteurs", spF, user.token, editVisiteur._spId);
      toast("Visiteur mis à jour dans SharePoint ✓");
    } else {
      if (!user?.token) {
        toast("⚠️ Non connecté à SharePoint — visiteur non enregistré");
      } else {
        saveToSP("ECOGEST_Visiteurs", spF, user.token).then((result) => {
          if (!result) {
            toast("⚠️ Échec de l'enregistrement dans SharePoint — visiteur non ajouté");
            return;
          }
          setData((d) => ({ ...d, visiteurs: [...(d.visiteurs || []), { ...newV, _spId: result.id }] }));
          toast("Visiteur enregistré dans SharePoint ✓");
        });
      }
    }
    setShowModal(false);
    setForm({
      nom: "",
      prenom: "",
      cin: "",
      dateNaissance: "",
      lieuNaissance: "",
      genre: "M",
      telephone: "",
      email: "",
      adresse: "",
      ville: "",
      quartier: "",
      bacAnnee: "",
      bacMention: "Passable",
      bacSerie: "",
      etablissement: "",
      filiereS: "",
      dateVisite: /* @__PURE__ */ new Date().toISOString().split("T")[0],
      source: "Visite directe",
      statut: "Visiteur",
      observations: "",
      recuPar: "",
      anneeUniversitaire: data.configuration?.anneeUniversitaireActive || ANNEES_UNIVERSITAIRES[0],
      documents: [],
      tarif: [],
    });
  };
  const handleUpdateFiche = (updated) => {
    const spF = {
      Title: ((updated.nom || "") + " " + (updated.prenom || "")).toUpperCase(),
      Nom: (updated.nom || "").toUpperCase(),
      Prenom: (updated.prenom || "").toUpperCase(),
      CIN: (updated.cin || "").toUpperCase(),
      DateNaissance: updated.dateNaissance || "",
      LieuNaissance: updated.lieuNaissance || "",
      Genre: updated.genre || "M",
      Telephone: updated.telephone || "",
      Email: updated.email || "",
      Adresse: updated.adresse || "",
      Ville: updated.ville || "",
      Quartier: updated.quartier || "",
      BacAnnee: updated.bacAnnee || "",
      BacMention: updated.bacMention || "",
      BacSerie: updated.bacSerie || "",
      Etablissement: updated.etablissement || "",
      FiliereS: updated.filiereS || "",
      DateVisite: updated.dateVisite || "",
      Source: updated.source || "Visite directe",
      Statut: updated.statut || "Visiteur",
      Observations: updated.observations || "",
      RecuPar: updated.recuPar || "",
    };
    let etudiantSync = null;
    setData((d) => {
      const linked = updated.cin ? (d.etudiants || []).find((e) => e.cin && e.cin === updated.cin) : null;
      let etudiants = d.etudiants || [];
      if (linked) {
        etudiantSync = {
          ...linked,
          nom: updated.nom || "",
          prenom: updated.prenom || "",
          cin: updated.cin || "",
          dateNaissance: updated.dateNaissance || "",
          lieuNaissance: updated.lieuNaissance || "",
          genre: updated.genre || "M",
          telephone: updated.telephone || "",
          email: updated.email || "",
          adresse: updated.adresse || "",
          ville: updated.ville || "",
          quartier: updated.quartier || "",
        };
        etudiants = etudiants.map((e) => (e.id === linked.id ? etudiantSync : e));
      }
      return { ...d, visiteurs: (d.visiteurs || []).map((v) => (v.id === updated.id ? { ...updated } : v)), etudiants };
    });
    if (user?.token && updated._spId) saveToSP("ECOGEST_Visiteurs", spF, user.token, updated._spId);
    if (etudiantSync && user?.token && etudiantSync._spId) {
      saveToSP(
        "ECOGEST_Etudiants",
        {
          Nom: (etudiantSync.nom || "").toUpperCase(),
          Prenom: (etudiantSync.prenom || "").toUpperCase(),
          CIN: (etudiantSync.cin || "").toUpperCase(),
          DateNaissance: etudiantSync.dateNaissance || "",
          LieuNaissance: etudiantSync.lieuNaissance || "",
          Genre: etudiantSync.genre || "M",
          Telephone: etudiantSync.telephone || "",
          Email: etudiantSync.email || "",
          Adresse: etudiantSync.adresse || "",
          Ville: etudiantSync.ville || "",
          Quartier: etudiantSync.quartier || "",
        },
        user.token,
        etudiantSync._spId
      );
    }
    setShowFiche(updated);
    toast(
      etudiantSync
        ? "Visiteur et fiche étudiant liée mis à jour dans SharePoint ✓"
        : "Visiteur mis à jour dans SharePoint ✓"
    );
  };
  const handleConvertir = (id, newStatut) => {
    const v = (data.visiteurs || []).find((x) => x.id === id);
    if (!v) return;
    const marquerStatutVisiteur = (statut) => {
      setData((d) => ({ ...d, visiteurs: (d.visiteurs || []).map((x) => (x.id === id ? { ...x, statut } : x)) }));
      if (user?.token && v._spId) saveToSP("ECOGEST_Visiteurs", { Statut: statut }, user.token, v._spId);
    };
    if (newStatut === "Inscrit") {
      const anneeActive = data.configuration?.anneeUniversitaireActive || ANNEES_UNIVERSITAIRES[0];
      const numInscription = genNumInscriptionESRMI(data.etudiants, anneeActive, data.paiements?.etudiants_paiement);
      const newEt = {
        id: Date.now(),
        nom: v.nom || "",
        prenom: v.prenom || "",
        cin: v.cin || "",
        dateNaissance: v.dateNaissance || "",
        lieuNaissance: v.lieuNaissance || "",
        filiere: v.filiereS || "",
        niveau: "BAC+3",
        annee: "",
        email: v.email || "",
        telephone: v.telephone || "",
        statut: "Actif",
        genre: v.genre || "M",
        ville: v.ville || "",
        quartier: v.quartier || "",
        adresse: v.adresse || "",
        numInscription,
        promotion: anneeActive,
        responsableFinancier: "",
      };
      if (!user?.token) {
        toast("⚠️ Non connecté à SharePoint — étudiant non créé");
      } else {
        const spEt = {
          Title: (newEt.nom + " " + newEt.prenom).toUpperCase(),
          Nom: (newEt.nom || "").toUpperCase(),
          Prenom: (newEt.prenom || "").toUpperCase(),
          CIN: (newEt.cin || "").toUpperCase(),
          DateNaissance: newEt.dateNaissance,
          LieuNaissance: newEt.lieuNaissance,
          Genre: newEt.genre,
          Email: newEt.email,
          Telephone: newEt.telephone,
          Adresse: newEt.adresse,
          Ville: newEt.ville,
          Quartier: newEt.quartier,
          Filiere: newEt.filiere,
          Niveau: newEt.niveau,
          Annee: newEt.annee,
          Promotion: newEt.promotion,
          Statut: newEt.statut,
          NumInscription: newEt.numInscription,
          ResponsableFinancier: newEt.responsableFinancier,
        };
        saveToSP("ECOGEST_Etudiants", spEt, user.token).then((result) => {
          if (!result) {
            toast("⚠️ Échec de la création de l'étudiant dans SharePoint");
            return;
          }
          setData((d) => ({ ...d, etudiants: [...(d.etudiants || []), { ...newEt, _spId: result.id }] }));
          marquerStatutVisiteur("Inscrit");
          toast("✓ Visiteur converti en Étudiant — enregistré dans SharePoint !");
        });
      }
    } else {
      marquerStatutVisiteur(newStatut);
      toast("Statut mis à jour dans SharePoint ✓");
    }
    setShowFiche(null);
  };
  const handleDeleteVisiteur = (v) => {
    if (!window.confirm(`Supprimer le visiteur "${v.prenom} ${v.nom}" ? Cette action est irréversible.`)) return;
    setData((d) => ({ ...d, visiteurs: (d.visiteurs || []).filter((x) => x.id !== v.id) }));
    if (user?.token && v._spId) deleteFromSP("ECOGEST_Visiteurs", v._spId, user.token);
    toast("Visiteur supprimé de SharePoint ✓");
    setShowFiche(null);
  };
  const stats = {
    total: (data.visiteurs || []).length,
    visiteurs: (data.visiteurs || []).filter((v) => v.statut === "Visiteur").length,
    inscrits: (data.visiteurs || []).filter((v) => v.statut === "Inscrit").length,
    nonInt: (data.visiteurs || []).filter((v) => v.statut === "Non intéressé").length,
  };
  return (
    /* @__PURE__ */ <div>
      <PageHeader
        title="Visiteurs & Prospects"
        sub={`${stats.total} visiteur(s) — liste SharePoint "${SP.lists.visiteurs}"`}
        action={canEdit && /* @__PURE__ */ <Btn onClick={() => setShowModal(true)}>+ Nouveau visiteur</Btn>}
      />
      <div style={{ display: "flex", gap: 12, marginBottom: 20 }}>
        <StatCard label="Total visiteurs" value={stats.total} color={C.navy} icon="👥" />
        <StatCard label="En attente" value={stats.visiteurs} color={C.orange} icon="⏳" />
        <StatCard label="Inscrits" value={stats.inscrits} color={C.green} icon="✅" />
        <StatCard label="Non intéressés" value={stats.nonInt} color={C.red} icon="❌" />
        <StatCard
          label="Taux conversion"
          value={stats.total ? `${Math.round((stats.inscrits / stats.total) * 100)}%` : "—"}
          color={C.purple}
          icon="📈"
        />
      </div>
      <div style={{ display: "flex", gap: 10, marginBottom: 14, flexWrap: "wrap" }}>
        <input
          placeholder="🔍 Rechercher nom, CIN, téléphone..."
          value={search}
          onChange={(e) => setSearch(e.target.value)}
          style={{
            flex: 1,
            minWidth: 200,
            padding: "9px 14px",
            border: `1px solid ${C.border}`,
            borderRadius: 8,
            fontSize: 13,
            fontFamily: "inherit",
          }}
        />
        {["", "Visiteur", "Inscrit", "Non intéressé"].map((s) => (
          /* @__PURE__ */ <button
            key={s}
            onClick={() => setFilterStatut(s)}
            style={{
              padding: "8px 14px",
              borderRadius: 20,
              border: `1px solid ${C.border}`,
              cursor: "pointer",
              background: filterStatut === s ? C.navy : C.white,
              color: filterStatut === s ? C.white : "#555",
              fontSize: 12,
              fontFamily: "inherit",
            }}
          >
            {s || "Tous"}
          </button>
        ))}
      </div>
      <Card style={{ overflow: "hidden" }}>
        <DataTable
          cols={[
            {
              key: "_filiereCode",
              label: "Filière souhaitée",
              render: (_, row) => (
                /* @__PURE__ */ <FiliereBadge code={filiereCode(row.filiereS, data.filieres) || row.filiereS || "—"} />
              ),
            },
            {
              key: "nom",
              label: "Nom",
              render: (v, row) => (
                /* @__PURE__ */ <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
                  <span>{v}</span>
                  {estEtudiantTest(row.prenom, row.nom) && (
                    /* @__PURE__ */ <span
                      title="Visiteur de test — n'existe pas réellement"
                      style={{
                        background: C.purple,
                        color: "white",
                        fontSize: 9.5,
                        fontWeight: 800,
                        borderRadius: 20,
                        padding: "1px 7px",
                        flexShrink: 0,
                        display: "inline-flex",
                        alignItems: "center",
                        gap: 3,
                      }}
                    >
                      🧪 TEST
                    </span>
                  )}
                </div>
              ),
            },
            { key: "prenom", label: "Prénom" },
            { key: "cin", label: "CIN" },
            { key: "telephone", label: "Téléphone" },
            { key: "ville", label: "Ville" },
            { key: "dateVisite", label: "Date visite", render: (v) => formatDateFR(v) },
            {
              key: "statut",
              label: "Statut",
              render: (v) => (
                /* @__PURE__ */ <Badge color={v === "Inscrit" ? C.green : v === "Non intéressé" ? C.red : C.orange}>
                  {v}
                </Badge>
              ),
            },
            {
              key: "_fiche",
              label: "Fiche",
              render: (_, row) => (
                /* @__PURE__ */ <Btn small={true} variant="navy" onClick={() => setShowFiche(row)}>
                  Voir fiche
                </Btn>
              ),
            },
            ...(canEdit
              ? [
                  {
                    key: "_suppr",
                    label: "Suppr.",
                    render: (_, row) => (
                      /* @__PURE__ */ <Btn small={true} variant="danger" onClick={() => handleDeleteVisiteur(row)}>
                        🗑
                      </Btn>
                    ),
                  },
                ]
              : []),
          ]}
          rows={filteredSortedVisiteurs}
          emptyMsg="Aucun visiteur enregistré"
          compact={true}
          rowStyle={(row) => ({
            background: bandeFiliereVisParId[row.id] === 0 ? C.navyLight : "#e0f4f2",
            borderTop: premiereLigneFiliereVisId[row.id] ? `2.5px solid ${C.navy}` : void 0,
            borderBottom: derniereLigneFiliereVisId[row.id] ? `2.5px solid ${C.navy}` : `1px solid ${C.border}`,
          })}
        />
      </Card>
      {showModal && (
        /* @__PURE__ */ <Modal title="Nouveau visiteur → SharePoint" onClose={() => setShowModal(false)} width={640}>
          <div
            style={{
              fontSize: 11,
              fontWeight: 700,
              color: C.orange,
              textTransform: "uppercase",
              letterSpacing: "0.08em",
              marginBottom: 10,
            }}
          >
            Informations personnelles
          </div>
          <div className="eco-grid3" style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: "0 14px" }}>
            <Input label="Nom *" value={form.nom} onChange={(e) => setForm((f) => ({ ...f, nom: e.target.value }))} />
            <Input
              label="Prénom *"
              value={form.prenom}
              onChange={(e) => setForm((f) => ({ ...f, prenom: e.target.value }))}
            />
            <Input label="CIN *" value={form.cin} onChange={(e) => setForm((f) => ({ ...f, cin: e.target.value }))} />
          </div>
          <div className="eco-grid2" style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "0 14px" }}>
            <Input
              label="Date de naissance"
              type="date"
              value={form.dateNaissance}
              onChange={(e) => setForm((f) => ({ ...f, dateNaissance: e.target.value }))}
            />
            <Input
              label="Lieu de naissance"
              value={form.lieuNaissance}
              onChange={(e) => setForm((f) => ({ ...f, lieuNaissance: e.target.value }))}
            />
          </div>
          <div className="eco-grid3" style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: "0 14px" }}>
            <Sel label="Genre" value={form.genre} onChange={(e) => setForm((f) => ({ ...f, genre: e.target.value }))}>
              <option value="M">Masculin</option>
              <option value="F">Féminin</option>
            </Sel>
            <div>
              <Input
                label="Téléphone *"
                value={form.telephone}
                onChange={(e) => setForm((f) => ({ ...f, telephone: e.target.value }))}
                placeholder="06XXXXXXXX"
                style={form.telephone && !/^\d{10}$/.test(form.telephone) ? { borderColor: C.red } : {}}
              />
              {form.telephone && !/^\d{10}$/.test(form.telephone) && (
                /* @__PURE__ */ <div style={{ fontSize: 11, color: C.red, marginTop: -10, marginBottom: 13 }}>
                  ⚠️ Le numéro doit contenir exactement 10 chiffres
                </div>
              )}
            </div>
            <Input
              label="Email"
              type="email"
              value={form.email}
              onChange={(e) => setForm((f) => ({ ...f, email: e.target.value }))}
            />
          </div>
          <div className="eco-grid3" style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: "0 14px" }}>
            <Input
              label="Adresse"
              value={form.adresse}
              onChange={(e) => setForm((f) => ({ ...f, adresse: e.target.value }))}
            />
            <Input
              label="Ville"
              value={form.ville}
              onChange={(e) => setForm((f) => ({ ...f, ville: e.target.value }))}
            />
            <Input
              label="Quartier"
              value={form.quartier}
              onChange={(e) => setForm((f) => ({ ...f, quartier: e.target.value }))}
            />
          </div>
          <div
            style={{
              fontSize: 11,
              fontWeight: 700,
              color: C.orange,
              textTransform: "uppercase",
              letterSpacing: "0.08em",
              margin: "10px 0",
            }}
          >
            Baccalauréat & Formation
          </div>
          <div className="eco-grid3" style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: "0 14px" }}>
            <Input
              label="Année du Bac"
              value={form.bacAnnee}
              onChange={(e) => setForm((f) => ({ ...f, bacAnnee: e.target.value }))}
              placeholder="2024"
            />
            <Sel
              label="Mention"
              value={form.bacMention}
              onChange={(e) => setForm((f) => ({ ...f, bacMention: e.target.value }))}
            >
              <option>Très bien</option>
              <option>Bien</option>
              <option>Assez bien</option>
              <option>Passable</option>
            </Sel>
            <Input
              label="Série du Bac"
              value={form.bacSerie}
              onChange={(e) => setForm((f) => ({ ...f, bacSerie: e.target.value }))}
              placeholder="Sciences Éco..."
            />
          </div>
          <div className="eco-grid3" style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: "0 14px" }}>
            <Input
              label="Établissement d'origine"
              value={form.etablissement}
              onChange={(e) => setForm((f) => ({ ...f, etablissement: e.target.value }))}
            />
            <Sel
              label="Filière souhaitée"
              value={form.filiereS}
              onChange={(e) => setForm((f) => ({ ...f, filiereS: e.target.value }))}
            >
              <option value="">-- Sélectionner une filière --</option>
              {filieresList.length === 0 && (
                /* @__PURE__ */ <option disabled={true}>Aucune filière disponible dans SharePoint</option>
              )}
              {filieresList.map((f) => (
                /* @__PURE__ */ <option key={f} value={f}>
                  {f}
                </option>
              ))}
            </Sel>
            <Input
              label="Date de visite"
              type="date"
              value={form.dateVisite}
              onChange={(e) => setForm((f) => ({ ...f, dateVisite: e.target.value }))}
            />
          </div>
          <div className="eco-grid3" style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: "0 14px" }}>
            <Sel
              label="Source"
              value={form.source}
              onChange={(e) => setForm((f) => ({ ...f, source: e.target.value }))}
            >
              <option>Visite directe</option>
              <option>Téléphone</option>
              <option>Email</option>
              <option>Réseaux sociaux</option>
              <option>Recommandation</option>
              <option>Salon</option>
            </Sel>
            <Sel
              label="Statut"
              value={form.statut}
              onChange={(e) => setForm((f) => ({ ...f, statut: e.target.value }))}
            >
              <option>Visiteur</option>
              <option>Prospect</option>
              <option>Inscrit</option>
              <option>Non intéressé</option>
            </Sel>
            <Sel
              label="Reçu par"
              value={form.recuPar}
              onChange={(e) => setForm((f) => ({ ...f, recuPar: e.target.value }))}
            >
              <option value="">-- Sélectionner --</option>
              <option>ADILA BAHALLA</option>
              <option>KAWTAR BOUABID</option>
              <option>FATIMA ZAHRA ZNIBER</option>
            </Sel>
          </div>
          <div style={{ display: "flex", gap: 10, justifyContent: "flex-end" }}>
            <Btn variant="light" onClick={() => setShowModal(false)}>
              Annuler
            </Btn>
            <Btn
              onClick={handleSave}
              disabled={!form.nom || !form.prenom || !form.cin || !/^\d{10}$/.test(form.telephone || "")}
            >
              Enregistrer → SharePoint
            </Btn>
          </div>
        </Modal>
      )}
      {showFiche && (
        /* @__PURE__ */ <FicheVisiteur
          visiteur={showFiche}
          onClose={() => setShowFiche(null)}
          onConvertir={handleConvertir}
          onSave={handleUpdateFiche}
          onDelete={handleDeleteVisiteur}
          canEdit={canEdit}
          user={user}
          filieresList={filieresList}
        />
      )}
    </div>
  );
}
const PAIEMENT_ADMINS = ["k.bouabid@esrmi.ma", "a.bahalla@esrmi.ma", "fz.zniber@esrmi.ma", "k.oubdi@esrmi.ma"];
const canAccessPaiement = (user) => user && PAIEMENT_ADMINS.includes(user.email);
const DEMO_PAIEMENT = {
  parametres: [
    {
      id: 1,
      filiere: "BAC+3 en Management et Gestion des entreprises",
      niveau: "BAC+3",
      annee: "2026-2027",
      fraisInscription: 2e3,
      fraisScolarite: 18e3,
      fraisAssurance: 300,
      fraisDossier: 500,
      fraisDiplomation: 1e3,
      fraisRattrapage: 500,
      fraisDivers: 200,
      total: 22500,
    },
    {
      id: 2,
      filiere: "BAC+3 en Systèmes d'Information – Option Génie Logiciel",
      niveau: "BAC+3",
      annee: "2026-2027",
      fraisInscription: 2e3,
      fraisScolarite: 19500,
      fraisAssurance: 300,
      fraisDossier: 500,
      fraisDiplomation: 1e3,
      fraisRattrapage: 500,
      fraisDivers: 200,
      total: 24e3,
    },
    {
      id: 3,
      filiere: "BAC+5 en Cybersécurité des réseaux et des systèmes d'information",
      niveau: "BAC+5",
      annee: "2026-2027",
      fraisInscription: 2500,
      fraisScolarite: 25e3,
      fraisAssurance: 400,
      fraisDossier: 600,
      fraisDiplomation: 1500,
      fraisRattrapage: 600,
      fraisDivers: 300,
      total: 30900,
    },
  ],
  etudiants_paiement: [
    {
      id: 1,
      etudiantId: 1,
      nom: "Alami",
      prenom: "Youssef",
      numInscription: "INS2026001",
      filiere: "BAC+3 en Management et Gestion des entreprises",
      niveau: "BAC+3",
      promotion: "2026-2027",
      statut: "Inscrit",
      cin: "BK123456",
      telephone: "0612345678",
      email: "y.alami@esr.ma",
      responsableFinancier: "",
      totalAPayer: 22500,
      totalPaye: 15e3,
      remises: [],
      versements: [
        {
          id: 1,
          date: "2026-09-10",
          montant: 7500,
          mode: "Virement",
          reference: "VIR001",
          banque: "CIH",
          numCheque: "",
          agent: "k.bouabid@esrmi.ma",
          recu: "REC-2026-001",
          anneeUniversitaire: "2026-2027",
          motif: "Frais d'inscription",
        },
        {
          id: 2,
          date: "2027-01-12",
          montant: 7500,
          mode: "Chèque",
          reference: "CHQ002",
          banque: "Attijariwafa",
          numCheque: "123456",
          agent: "a.bahalla@esrmi.ma",
          recu: "REC-2027-001",
          anneeUniversitaire: "2026-2027",
          motif: "Frais de formation",
        },
      ],
      penalites: [],
      historique: [],
    },
    {
      id: 2,
      etudiantId: 2,
      nom: "Benali",
      prenom: "Sara",
      numInscription: "INS2026002",
      filiere: "BAC+3 en Systèmes d'Information – Option Génie Logiciel",
      niveau: "BAC+3",
      promotion: "2026-2027",
      statut: "Inscrit",
      cin: "BE234567",
      telephone: "0623456789",
      email: "s.benali@esr.ma",
      responsableFinancier: "Hassan Benali (Père) - 0661234567",
      totalAPayer: 24e3,
      totalPaye: 0,
      remises: [{ id: 1, type: "Bourse", pourcentage: 20, montant: 4800, motif: "Bourse d'excellence", valide: true }],
      versements: [],
      penalites: [{ id: 1, date: "2026-10-01", montant: 375, motif: "Retard > 15 jours", applique: true }],
      historique: [],
    },
  ],
  parametres_penalites: { actif: true, delaiJours: 15, typeCalcul: "pourcentage", valeur: 5 },
  parametres_alertes: { email: true, sms: false, rappel1: 7, rappel2: 3, rappel3: 0 },
};
const formatMAD = (n) => `${(n || 0).toLocaleString("fr-MA")} MAD`;
const today = () => /* @__PURE__ */ new Date().toISOString().split("T")[0];
const formatDateFR = (iso) => {
  if (!iso) return "";
  const [y, m, d] = iso.split("-");
  return y && m && d ? `${d}/${m}/${y}` : iso;
};
const genRecu = () =>
  `REC-${/* @__PURE__ */ new Date().getFullYear()}-${String(Math.floor(Math.random() * 9e3) + 1e3)}`;
const genNumInscriptionESRMI = (etudiants, year, paiements) => {
  const anneeRef = year || /* @__PURE__ */ new Date().getFullYear();
  const prefix = `ESRMI-${anneeRef}/`;
  const numerosUtilises = [
    ...(etudiants || []).map((e) => e.numInscription || ""),
    ...(paiements || []).map((p) => p.numInscription || ""),
  ];
  const nums = numerosUtilises
    .filter((n) => n.startsWith(prefix))
    .map((n) => parseInt(n.slice(prefix.length), 10))
    .filter((n) => !isNaN(n));
  const next = nums.length ? Math.max(...nums) + 1 : 1;
  return `${prefix}${next}`;
};
const ANNEES_UNIVERSITAIRES = ["2025-2026", "2026-2027", "2027-2028", "2028-2029", "2029-2030"];
const MOTIFS_PAIEMENT = [
  "Frais de formation",
  "Frais d'inscription",
  "Frais d'examen",
  "Frais de rattrapage",
  "Activités parascolaires",
  "Autre",
];
function SoldeBadge({ ep }) {
  const restant = ep.totalAPayer - ep.totalPaye;
  const taux = ep.totalAPayer > 0 ? Math.round((ep.totalPaye / ep.totalAPayer) * 100) : 0;
  const color = taux >= 100 ? C.green : taux >= 50 ? C.orange : C.red;
  return (
    /* @__PURE__ */ <div style={{ minWidth: 230 }}>
      <div style={{ display: "flex", gap: 8, flexWrap: "wrap", marginBottom: 7 }}>
        <span
          style={{
            background: C.navy + "18",
            color: C.navy,
            border: `1px solid ${C.navy}40`,
            borderRadius: 20,
            padding: "2px 8px",
            fontSize: 11,
            fontWeight: 600,
          }}
        >
          Tarif formation: {formatMAD(ep.totalAPayer)}
        </span>
        <span
          style={{
            background: C.green + "18",
            color: C.green,
            border: `1px solid ${C.green}40`,
            borderRadius: 20,
            padding: "2px 8px",
            fontSize: 11,
            fontWeight: 600,
          }}
        >
          Payé: {formatMAD(ep.totalPaye)}
        </span>
        <span
          style={{
            background: color + "18",
            color,
            border: `1px solid ${color}40`,
            borderRadius: 20,
            padding: "2px 8px",
            fontSize: 11,
            fontWeight: 600,
          }}
        >
          Reste: {formatMAD(restant)}
        </span>
      </div>
      <div style={{ display: "flex", alignItems: "center", gap: 9 }}>
        <div
          style={{
            flex: 1,
            background: "#e7ebf2",
            borderRadius: 20,
            height: 14,
            boxShadow: "inset 0 1px 4px rgba(20,30,60,0.22)",
            overflow: "hidden",
          }}
        >
          <div
            style={{
              width: `${Math.min(taux, 100)}%`,
              height: "100%",
              background: `linear-gradient(90deg, ${color}bb 0%, ${color} 100%)`,
              borderRadius: 20,
              transition: "width 0.5s",
              boxShadow: `0 0 8px ${color}99`,
            }}
          />
        </div>
        <span style={{ fontSize: 13, fontWeight: 800, color, minWidth: 36, textAlign: "right" }}>{taux}%</span>
      </div>
    </div>
  );
}
function genererBilletEntree(p) {
  const html = `
    <html><head><title>Billet d'entrée</title>
    <style>
      body { font-family: Arial, sans-serif; padding: 40px; max-width: 600px; margin: 0 auto; }
      .header { text-align: center; border-bottom: 2px solid #1a2f5e; padding-bottom: 20px; margin-bottom: 20px; }
      .sub { font-size: 16px; font-weight: 600; color: #1a2f5e; }
      .titre { font-size: 20px; font-weight: bold; text-align: center; margin: 20px 0; color: #e56b2d; }
      .row { display: flex; justify-content: space-between; padding: 8px 0; border-bottom: 1px solid #f0f0f0; }
      .label { color: #888; font-size: 13px; }
      .value { font-weight: bold; font-size: 13px; color: #333; }
      .ok { text-align: center; padding: 16px; background: #eaf7ef; color: #2e7d52; border-radius: 8px; margin: 20px 0; font-weight: bold; font-size: 15px; }
      .footer { text-align: center; margin-top: 30px; font-size: 11px; color: #aaa; border-top: 1px solid #eee; padding-top: 15px; }
      .signature { margin-top: 40px; display: flex; justify-content: space-between; }
    </style></head><body>
    <div class="header">
      <img src="${LOGO}" alt="ESRMI" style="width:175px;margin-bottom:10px;" />
      <div class="sub">École Supérieure de Rabat en Management et Ingénierie</div>
    </div>
    <div class="titre">BILLET D'ENTRÉE</div>
    <div class="row"><span class="label">Étudiant</span><span class="value">${p.etudiant}</span></div>
    <div class="row"><span class="label">N° Inscription</span><span class="value">${p.numInscription || "—"}</span></div>
    <div class="row"><span class="label">Filière</span><span class="value">${p.filiere || "—"}</span></div>
    <div class="row"><span class="label">Date de l'absence</span><span class="value">${formatDateFR(p.date)}</span></div>
    <div class="row"><span class="label">Module / Séance</span><span class="value">${p.module || "—"} · ${p.seance || p.heure || "—"}</span></div>
    <div class="row"><span class="label">Document de justification</span><span class="value">${p.documentJustification || "—"}</span></div>
    <div class="row"><span class="label">Justifié par</span><span class="value">${p.justifiePar || "—"}</span></div>
    <div class="row"><span class="label">Date de justification</span><span class="value">${formatDateFR(p.dateJustification) || "—"}</span></div>
    <div class="ok">✓ Absence justifiée — Entrée en cours autorisée</div>
    <div class="signature">
      <div style="text-align:center"><div style="border-top:1px solid #333;width:150px;margin:40px 0 5px">Cachet & Signature</div></div>
    </div>
    <div class="footer">Document généré le ${/* @__PURE__ */ new Date().toLocaleString("fr-FR")} © ${/* @__PURE__ */ new Date().getFullYear()}</div>
    <div class="no-print" style="text-align:center;margin-top:24px;">
      <button onclick="window.print()" style="background:#1a2f5e;color:white;border:none;padding:10px 24px;border-radius:8px;font-size:14px;cursor:pointer;font-family:inherit;">🖨️ Imprimer le billet</button>
    </div>
    <style>@media print { .no-print { display:none; } }</style>
    </body></html>`;
  const w = window.open("", "_blank");
  if (!w) {
    alert("Le billet n'a pas pu s'ouvrir : autorisez les pop-ups pour ce site.");
    return;
  }
  w.document.write(html);
  w.document.close();
}
function genererRecu(ep, versement) {
  const html = `
    <html><head><title>Reçu ${versement.recu}</title>
    <style>
      body { font-family: Arial, sans-serif; padding: 40px; max-width: 600px; margin: 0 auto; }
      .header { text-align: center; border-bottom: 2px solid #1a2f5e; padding-bottom: 20px; margin-bottom: 20px; }
      .logo-text { font-size: 24px; font-weight: bold; color: #1a2f5e; }
      .sub { font-size: 12px; color: #888; }
      .recu-title { font-size: 20px; font-weight: bold; text-align: center; margin: 20px 0; color: #e56b2d; }
      .row { display: flex; justify-content: space-between; padding: 8px 0; border-bottom: 1px solid #f0f0f0; }
      .label { color: #888; font-size: 13px; }
      .value { font-weight: bold; font-size: 13px; color: #333; }
      .montant { font-size: 24px; font-weight: bold; color: #1a2f5e; text-align: center; padding: 20px; background: #f5f5f3; border-radius: 8px; margin: 20px 0; }
      .footer { text-align: center; margin-top: 30px; font-size: 11px; color: #aaa; border-top: 1px solid #eee; padding-top: 15px; }
      .signature { margin-top: 40px; display: flex; justify-content: space-between; }
    </style></head><body>
    <div class="header">
      <img src="${LOGO}" alt="ESRMI" style="width:175px;margin-bottom:10px;" />
      <div class="sub" style="font-size:16px;font-weight:600;color:#1a2f5e;">École Supérieure de Rabat en Management et Ingénierie</div>
    </div>
    <div class="recu-title">REÇU DE PAIEMENT — ${versement.recu}</div>
    <div class="row"><span class="label">Étudiant</span><span class="value">${ep.prenom} ${ep.nom}</span></div>
    <div class="row"><span class="label">N° Inscription</span><span class="value">${ep.numInscription}</span></div>
    <div class="row"><span class="label">Filière</span><span class="value">${ep.filiere}</span></div>
    <div class="row"><span class="label">Date de paiement</span><span class="value">${formatDateFR(versement.date)}</span></div>
    <div class="row"><span class="label">Année universitaire</span><span class="value">${versement.anneeUniversitaire || "—"}</span></div>
    <div class="row"><span class="label">Motif</span><span class="value">${versement.motif || "—"}</span></div>
    <div class="row"><span class="label">Mode de paiement</span><span class="value">${versement.mode}</span></div>
    ${versement.banque ? `<div class="row"><span class="label">Banque</span><span class="value">${versement.banque}</span></div>` : ""}
    ${versement.numCheque ? `<div class="row"><span class="label">N° Chèque</span><span class="value">${versement.numCheque}</span></div>` : ""}
    <div class="row"><span class="label">Référence</span><span class="value">${versement.reference}</span></div>
    <div class="row"><span class="label">Opérateur</span><span class="value">${versement.agent}</span></div>
    <div class="montant">${formatMAD(versement.montant)}</div>
    <div class="signature">
      <div style="text-align:center"><div style="border-top:1px solid #333;width:150px;margin:40px 0 5px">Cachet & Signature</div></div>
    </div>
    <div class="footer">Document généré le ${/* @__PURE__ */ new Date().toLocaleString("fr-FR")} © ${/* @__PURE__ */ new Date().getFullYear()}</div>
    <div class="no-print" style="text-align:center;margin-top:24px;">
      <button onclick="window.print()" style="background:#1a2f5e;color:white;border:none;padding:10px 24px;border-radius:8px;font-size:14px;cursor:pointer;font-family:inherit;">🖨️ Imprimer le reçu</button>
    </div>
    <style>@media print { .no-print { display:none; } }</style>
    </body></html>`;
  const w = window.open("", "_blank");
  if (!w) {
    alert("Le reçu n'a pas pu s'ouvrir : autorisez les pop-ups pour ce site.");
    return;
  }
  w.document.write(html);
  w.document.close();
}
const NAVY_RGB = [0.1, 0.18, 0.37];
const ORANGE_RGB = [0.9, 0.42, 0.18];
const UNICODE_TO_WINANSI = {
  8364: 128,
  8218: 130,
  402: 131,
  8222: 132,
  8230: 133,
  8224: 134,
  8225: 135,
  710: 136,
  8240: 137,
  352: 138,
  8249: 139,
  338: 140,
  381: 142,
  8216: 145,
  8217: 146,
  8220: 147,
  8221: 148,
  8226: 149,
  8211: 150,
  8212: 151,
  732: 152,
  8482: 153,
  353: 154,
  8250: 155,
  339: 156,
  382: 158,
  376: 159,
};
function pdfEscapeText(str) {
  let out = "";
  for (const ch of String(str)) {
    let code = ch.codePointAt(0);
    if (code > 255) {
      code = UNICODE_TO_WINANSI[code];
      if (code === void 0) {
        out += "?";
        continue;
      }
    }
    const mapped = String.fromCharCode(code);
    if (mapped === "(" || mapped === ")" || mapped === "\\") out += "\\" + mapped;
    else out += mapped;
  }
  return out;
}
function base64ToBytes(b64) {
  const bin = atob(b64);
  const bytes = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
  return bytes;
}
function strToLatin1Bytes(s) {
  const arr = new Uint8Array(s.length);
  for (let i = 0; i < s.length; i++) arr[i] = s.charCodeAt(i) & 255;
  return arr;
}
const LOGO_ESR_B64 =
  "/9j/4AAQSkZJRgABAQEAeAB4AAD/2wBDAA0JCgsKCA0LCwsPDg0QFCEVFBISFCgdHhghMCoyMS8qLi00O0tANDhHOS0uQllCR05QVFVUMz9dY1xSYktTVFH/2wBDAQ4PDxQRFCcVFSdRNi42UVFRUVFRUVFRUVFRUVFRUVFRUVFRUVFRUVFRUVFRUVFRUVFRUVFRUVFRUVFRUVFRUVH/wAARCABSAGEDASIAAhEBAxEB/8QAHwAAAQUBAQEBAQEAAAAAAAAAAAECAwQFBgcICQoL/8QAtRAAAgEDAwIEAwUFBAQAAAF9AQIDAAQRBRIhMUEGE1FhByJxFDKBkaEII0KxwRVS0fAkM2JyggkKFhcYGRolJicoKSo0NTY3ODk6Q0RFRkdISUpTVFVWV1hZWmNkZWZnaGlqc3R1dnd4eXqDhIWGh4iJipKTlJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXGx8jJytLT1NXW19jZ2uHi4+Tl5ufo6erx8vP09fb3+Pn6/8QAHwEAAwEBAQEBAQEBAQAAAAAAAAECAwQFBgcICQoL/8QAtREAAgECBAQDBAcFBAQAAQJ3AAECAxEEBSExBhJBUQdhcRMiMoEIFEKRobHBCSMzUvAVYnLRChYkNOEl8RcYGRomJygpKjU2Nzg5OkNERUZHSElKU1RVVldYWVpjZGVmZ2hpanN0dXZ3eHl6goOEhYaHiImKkpOUlZaXmJmaoqOkpaanqKmqsrO0tba3uLm6wsPExcbHyMnK0tPU1dbX2Nna4uPk5ebn6Onq8vP09fb3+Pn6/9oADAMBAAIRAxEAPwD06iiigAquJtl6YG/jXen9R/X8asVn6yrJbx3Uf34HDfgeDWdRuMeZdDSmlKXK+poUUyCVZ4UlT7rDNPq076ohq2jCiiimIKKKKACiiigAooooAK5TxT4xj0mQ2dkqzXY+8W+7H9fU+1bHiPVBo+i3F5x5gG2MHu54H+P4V53oPhW+11zeXMjRQyEsZG5aTPUj/GtqcY/FLYlvsVpPFeuzSb21GRfZAFH5AVctvGWrpG0VxIl1GwwRIoB/AjH65rsbXwTokCAPA8zf3nc/04p83gzQ5FwLVoz6pIa0dSk1ZojlmtUyv4S8QWt5GbRn8qbOURz19QD3rqK8zufC5F00dlcEMHwok+vHIrb0nX77S5k0/wAQROgPCXDc/me49/zrmpxpSj+5d7dOprUc4yvVW52NFIrBlDKQVIyCO9LSAKKKKACiiigAooooA57XrD+2dXs7KUE2lspuZ1/vk8Iv6N+Fb0MSxRqigAAdqFjVXdwPmcjJ+lPpt30AKKKiupRDbu7HgClewbnMyNi9kYdnJ/WuhvPsM8bW92YGU9UkIrnLdTNOq93b+dc9qGgWfiP4qanZXrzJElskoMTAHIWMdwfU15+BbblI7cWrcqO306AadIYbO4FzZdfK37nh+nqvt1+tX7rULKyKi7vILcv90SyBd30yaxvDng7TPDd1Lc2Mty7ypsYSupGM57AelcZoWiWXjDXtbn1u5lN1FMUSBX2lFBI/IdP59a9Fu+rOFKx6nuUpv3DbjOc8Ypkc8MpxHKjn0Vga56x0abQvBd7YTXrXYSGYozDGxdpwo/z3rgdG8HwXfgh/EEF7cW9/Cskq7WGz5CenGQcDrmkM9jo61ymgatd6p4b0u9miM7sp81dufM2llZsevyg/VvpW3ppmWR0ljCHncqrhcjHIHud35UAaFFFFABRRRQA13VBljiuf1nURM32eM/KPvf4Vp3dte3OY0kit4z1ZQWY/yxTLPRbW1IYgyuO79PyrnrKdRckdF3Oml7On783d9ivoliykXMq44+QH+dcRquhf8JB8UtTs/tklrstkk3xjJOFjGOvv+leo151q8HiXTfH9/rGkaR9rjmhWIM/3SNqZ6Ed1q6VJUo8qMqlR1JczOh8LeFP+EdnuJf7Smu/OULiQY24P1ql4h8E2msXT6ppV4bLUgc+ZE3ysw9ccg+4/I1b8Naj4j1K4uIdd0iOzt/L+VlB+Y5xj7x7VzFknizwY11pen6V/aNnJIXt5Qpbbn1x06DIOOa1MzW8N63e6r4U1q01P5r2wSSGR/wC98p6++QRXBaZPrMmmabo82pfY9G1B2RXCAjO7DAng9ccZxzXofhrw7e6Z4X1M3vz6lqCySSKCDglThc9M5J/OqWm+Ep774brpN9AYL6N3kh34yr7iR+BHH40AdjYaXa6fpltYQKwit1Coc4b3OR696tRxLHnbnJ6kkkn8TWN4Qn1SXQootYtZILyD92xfH7wDo3+PuK3KACiiigCj/aFrJvePUYFRH8skkYDemSaeJlMjxC/hMiEBl4ypPTIzxmues9Ot1t7ZIdUQkhhEZYHYMrYU/K7EZ+ZcYwOvBFSTaZbrcsBqUaPbyu8YMBJUl0nbfgjd0GMY4Pc0AbqzK87wLfRNKgy0YwWUe4z70kdzHKgePUYHU/xKVI649fXisuzt4Ibq9mfUY5PtCbpFCOApZAcgFiuMKT0z71Ui0i3gQMdTLRqBJIRCzLsGxwASTg/Jnkng8ADFAG0+oWyKrHVLfaz7AQQct6daeLuEqzDUrchDhjleDgnB59AfyrAXSdNtVtGgvzHNAi72kEuZARgcBgQPkbjp7UxtBsJYzFHqUgaRdwDxsW6SpwCcgZkPA7j3oA6RLmN4/MTUYGT+8CpHfvn2P5Gmi8gIQjU7YiQ7Uwy/MfQc89RWJNpcMtvFLJqBEVwqx/6iQlyPMKkbmLD75znOcDGKda6HBBdpc2984WCYqRKkmSW8sYLbgW+4AM5GCOuKANprmJHkR9RgVohlwSoKD35460G5jVI3OowBZf8AVklcP9OeawrvTLTUpJZ01EFJpPNjEcLEqcR5yQc4+VfTBPWpLyzstQt4Emv2RoEliDbXCucqxzliWGEIIJOcn2oA2pbhIVZptQhjVW2EvtADYzjr1xTmuYbfLXF7CqgA/MQuM9D174Nc1Po1q80rf2tKXaXz5QYmxuO9eChUj/WY6k8AdBinppFnuP8AxNZHe4hFvgwDBUBQvAGQRt78cngZoA6H+1dN/wCgha/9/l/xoqz5af3F/KigDktL+fTLFn+YpLHtJ5x+8jHH5mrd2iPrkwZVYGaBTkdQcgj8aKKAI/Kj8+JPLXb9raPGONuUG36Y4xVO9Gy8niX5Y1tZyFHABAkUcfQAfQCiigCQ/vFgZ/mb7KxyeT95f8asaafMSzd/mYSR4J5PWWiigCbVmZNG0koSpEicjj/lm1aTgeTL/wBfaf8Aoa0UUAc5cAJDGEAUNcIpA4yDDESPxIBqSH57W9V/mCtJgHnHzLRRQBLGc600Z5RtuV7H956U2YmO6tQhK4nKDbxhQ2APpgUUUAdZRRRQB//Z";
function jpegDimensions(bytes) {
  let i = 2;
  while (i < bytes.length) {
    if (bytes[i] !== 255) {
      i++;
      continue;
    }
    const marker = bytes[i + 1];
    if (marker >= 192 && marker <= 195) {
      return { height: (bytes[i + 5] << 8) | bytes[i + 6], width: (bytes[i + 7] << 8) | bytes[i + 8] };
    }
    const len = (bytes[i + 2] << 8) | bytes[i + 3];
    i += 2 + len;
  }
  return { width: 250, height: 250 };
}
let _logoJpegBytes = null;
function getLogoJpegBytes() {
  if (!_logoJpegBytes) _logoJpegBytes = base64ToBytes(LOGO_JPEG_B64);
  return _logoJpegBytes;
}
let _measureCanvas = null;
function measureTextWidth(str, size, bold) {
  if (!_measureCanvas) _measureCanvas = document.createElement("canvas");
  const ctx = _measureCanvas.getContext("2d");
  ctx.font = `${bold ? "bold " : ""}${size}px Helvetica, Arial, sans-serif`;
  return ctx.measureText(str).width;
}
function wrapText(text, maxWidth, size, bold) {
  const words = text.split(" ");
  const lines = [];
  let current = "";
  for (const word of words) {
    const test = current ? current + " " + word : word;
    if (measureTextWidth(test, size, bold) > maxWidth && current) {
      lines.push(current);
      current = word;
    } else current = test;
  }
  if (current) lines.push(current);
  return lines;
}
class MiniPDF {
  constructor(w, h) {
    this.w = w;
    this.h = h;
    this.pages = [[]];
    this.ops = this.pages[0];
    this.images = [];
  } // Nouvelle page : les documents courts (attestations...) n'en ont jamais
  // besoin, mais un relevé d'absences peut dépasser une page.
  newPage() {
    this.pages.push([]);
    this.ops = this.pages[this.pages.length - 1];
  }
  text(x, y, str, opts = {}) {
    const { size = 11, bold = false, color = [0, 0, 0], angle = 0 } = opts;
    const font = bold ? "F2" : "F1";
    const [r, g, b] = color;
    if (angle) {
      const rad = (angle * Math.PI) / 180;
      const cosA = Math.cos(rad),
        sinA = Math.sin(rad);
      this.ops.push(
        `q ${r} ${g} ${b} rg BT /${font} ${size} Tf ${cosA.toFixed(5)} ${sinA.toFixed(5)} ${(-sinA).toFixed(5)} ${cosA.toFixed(5)} ${x.toFixed(2)} ${y.toFixed(2)} Tm (${pdfEscapeText(str)}) Tj ET Q`
      );
    } else {
      this.ops.push(
        `q ${r} ${g} ${b} rg BT /${font} ${size} Tf ${x.toFixed(2)} ${y.toFixed(2)} Td (${pdfEscapeText(str)}) Tj ET Q`
      );
    }
  }
  textCentered(cx, y, str, opts = {}) {
    this.text(cx - measureTextWidth(str, opts.size || 11, opts.bold) / 2, y, str, opts);
  }
  textRight(rx, y, str, opts = {}) {
    this.text(rx - measureTextWidth(str, opts.size || 11, opts.bold), y, str, opts);
  }
  paragraph(x, y, maxWidth, text, opts = {}) {
    const lines = wrapText(text, maxWidth, opts.size || 11, opts.bold);
    const leading = (opts.size || 11) * (opts.leading || 1.7);
    lines.forEach((l, i) => this.text(x, y - i * leading, l, opts));
    return y - lines.length * leading;
  }
  line(x1, y1, x2, y2, lw = 0.75, color = [0, 0, 0]) {
    const [r, g, b] = color;
    this.ops.push(
      `q ${r} ${g} ${b} RG ${lw} w ${x1.toFixed(2)} ${y1.toFixed(2)} m ${x2.toFixed(2)} ${y2.toFixed(2)} l S Q`
    );
  }
  rect(x, y, w, h, fill) {
    const [r, g, b] = fill;
    this.ops.push(`q ${r} ${g} ${b} rg ${x.toFixed(2)} ${y.toFixed(2)} ${w.toFixed(2)} ${h.toFixed(2)} re f Q`);
  }
  image(jpegBytes, x, y, w, h) {
    const name = "Im" + (this.images.length + 1);
    this.images.push({ name, bytes: jpegBytes });
    this.ops.push(`q ${w.toFixed(2)} 0 0 ${h.toFixed(2)} ${x.toFixed(2)} ${y.toFixed(2)} cm /${name} Do Q`);
  }
  build() {
    const parts = [];
    let offset = 0;
    const offsets = {};
    const push = (data) => {
      const b = typeof data === "string" ? strToLatin1Bytes(data) : data;
      parts.push(b);
      offset += b.length;
    };
    const N = this.pages.length;
    const fontF1Num = 3 + N * 2;
    const fontF2Num = fontF1Num + 1;
    const imgObjNums = {};
    let nextNum = fontF2Num + 1;
    for (const img of this.images) imgObjNums[img.name] = nextNum++;
    const xobjectDict = this.images.map((img) => `/${img.name} ${imgObjNums[img.name]} 0 R`).join(" ");
    const objs = {};
    const pageRefs = [];
    for (let i = 0; i < N; i++) {
      const pageNum = 3 + i * 2;
      const contentsNum = 4 + i * 2;
      pageRefs.push(`${pageNum} 0 R`);
      objs[pageNum] =
        `<< /Type /Page /Parent 2 0 R /MediaBox [0 0 ${this.w.toFixed(2)} ${this.h.toFixed(2)}] /Resources << /Font << /F1 ${fontF1Num} 0 R /F2 ${fontF2Num} 0 R >> /XObject << ${xobjectDict} >> >> /Contents ${contentsNum} 0 R >>`;
      objs[contentsNum] = { stream: strToLatin1Bytes(this.pages[i].join("\n")) };
    }
    objs[1] = `<< /Type /Catalog /Pages 2 0 R >>`;
    objs[2] = `<< /Type /Pages /Kids [${pageRefs.join(" ")}] /Count ${N} >>`;
    objs[fontF1Num] = `<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>`;
    objs[fontF2Num] = `<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold /Encoding /WinAnsiEncoding >>`;
    for (const img of this.images) {
      const { width, height } = jpegDimensions(img.bytes);
      objs[imgObjNums[img.name]] = {
        dict: `<< /Type /XObject /Subtype /Image /Width ${width} /Height ${height} /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode /Length ${img.bytes.length} >>`,
        stream: img.bytes,
      };
    }
    push("%PDF-1.4\n%âãÏÓ\n");
    const maxObj = Math.max(...Object.keys(objs).map(Number));
    for (let i = 1; i <= maxObj; i++) {
      offsets[i] = offset;
      const o = objs[i];
      if (o === void 0) continue;
      if (typeof o === "string")
        push(`${i} 0 obj
${o}
endobj
`);
      else if (o.stream && !o.dict) {
        push(`${i} 0 obj
<< /Length ${o.stream.length} >>
stream
`);
        push(o.stream);
        push(`
endstream
endobj
`);
      } else if (o.dict) {
        push(`${i} 0 obj
${o.dict}
stream
`);
        push(o.stream);
        push(`
endstream
endobj
`);
      }
    }
    const xrefOffset = offset;
    let xref = `xref
0 ${maxObj + 1}
0000000000 65535 f 
`;
    for (let i = 1; i <= maxObj; i++) {
      const off = offsets[i] !== void 0 ? offsets[i] : 0;
      xref += String(off).padStart(10, "0") + " 00000 n \n";
    }
    push(xref);
    push(`trailer
<< /Size ${maxObj + 1} /Root 1 0 R >>
startxref
${xrefOffset}
%%EOF`);
    const total = parts.reduce((s, p) => s + p.length, 0);
    const out = new Uint8Array(total);
    let pos = 0;
    for (const p of parts) {
      out.set(p, pos);
      pos += p.length;
    }
    return out;
  }
}
function downloadPDF(bytes, filename) {
  const blob = new Blob([bytes], { type: "application/pdf" });
  const url = URL.createObjectURL(blob);
  const fenetre = window.open(url, "_blank");
  if (!fenetre) {
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
  }
  setTimeout(() => URL.revokeObjectURL(url), 6e4);
}
const MM = 2.83465;
function genererAttestation(etudiant, filiereNom, type) {
  const DIRECTRICE = "Fatima Zahra ZNIBER";
  const anneeUniv = etudiant.promotion || ANNEES_UNIVERSITAIRES[0];
  const W = 210 * MM,
    H = 297 * MM;
  const titres = {
    inscription: "ATTESTATION D'INSCRIPTION",
    scolarite: "ATTESTATION DE SCOLARITÉ",
    reussite: "ATTESTATION DE RÉUSSITE",
  };
  const corps = {
    inscription: `est régulièrement inscrit(e) en ${filiereNom}, au sein de notre établissement, pour l'année universitaire ${anneeUniv}.`,
    scolarite: `poursuit sa scolarité en ${filiereNom} au sein de notre établissement, au titre de l'année universitaire ${anneeUniv}.`,
    reussite: `a suivi avec succès sa formation en ${filiereNom} au titre de l'année universitaire ${anneeUniv} et a satisfait aux exigences académiques requises pour la validation de cette année.`,
  };
  const doc = new MiniPDF(W, H);
  doc.rect(0, 0, W, 1.2 * MM, ORANGE_RGB);
  const HEADER_RESERVE = 65 * MM;
  const titreY = H - HEADER_RESERVE;
  doc.textCentered(W / 2, titreY, titres[type], { size: 19, bold: true, color: NAVY_RGB });
  doc.line(W / 2 - 42 * MM, titreY - 5 * MM, W / 2 + 42 * MM, titreY - 5 * MM, 1.2, ORANGE_RGB);
  const marge = 28 * MM;
  let y = titreY - 25 * MM;
  y = doc.paragraph(
    marge,
    y,
    W - 2 * marge,
    `Je soussignée, ${DIRECTRICE}, Directrice de l'École Supérieure de Rabat en Management et Ingénierie, atteste par la présente que :`,
    { size: 12, leading: 1.7 }
  );
  y -= 8 * MM;
  const idLine = `L'étudiant(e) ${etudiant.prenom} ${etudiant.nom}${etudiant.cin ? `, titulaire de la CIN n° ${etudiant.cin}` : ""}, ${corps[type]}`;
  y = doc.paragraph(marge, y, W - 2 * marge, idLine, { size: 12, leading: 1.7 });
  y -= 8 * MM;
  y = doc.paragraph(
    marge,
    y,
    W - 2 * marge,
    "La présente attestation est délivrée à l'intéressé(e) pour servir et valoir ce que de droit.",
    { size: 12, leading: 1.7 }
  );
  doc.text(W - 80 * MM, 70 * MM, `Fait à Rabat, le ${/* @__PURE__ */ new Date().toLocaleDateString("fr-FR")}`, {
    size: 10.5,
  });
  doc.text(W - 80 * MM, 55 * MM, "La Directrice", { size: 12, bold: true, color: NAVY_RGB });
  doc.text(W - 80 * MM, 48 * MM, DIRECTRICE, { size: 11 });
  doc.textCentered(W / 2, 8 * MM, "ESRMI — École Supérieure de Rabat en Management et Ingénierie", {
    size: 8,
    color: [0.5, 0.5, 0.5],
  });
  const noms = {
    inscription: "attestation_inscription",
    scolarite: "attestation_scolarite",
    reussite: "attestation_reussite",
  };
  downloadPDF(doc.build(), `${noms[type]}_${etudiant.prenom || ""}_${etudiant.nom || ""}.pdf`.replace(/\s+/g, "_"));
}
function genererReleveAbsences(etudiant, presencesEtudiant, coursListe = []) {
  const W = 210 * MM,
    H = 297 * MM;
  const marge = 20 * MM;
  const doc = new MiniPDF(W, H);
  const drawHeader = (premierePage) => {
    doc.rect(0, H - 1.2 * MM, W, 1.2 * MM, ORANGE_RGB);
    doc.textCentered(W / 2, H - 15 * MM, "ÉCOLE SUPÉRIEURE DE RABAT EN MANAGEMENT ET INGÉNIERIE", {
      size: 11,
      bold: true,
      color: NAVY_RGB,
    });
    if (premierePage) {
      doc.textCentered(W / 2, H - 29 * MM, "RELEVÉ D'ABSENCES", { size: 20, bold: true, color: NAVY_RGB });
      doc.line(W / 2 - 40 * MM, H - 33 * MM, W / 2 + 40 * MM, H - 33 * MM, 1.2, ORANGE_RGB);
    } else {
      doc.textCentered(W / 2, H - 25 * MM, "RELEVÉ D'ABSENCES (suite)", { size: 13, bold: true, color: NAVY_RGB });
    }
  };
  drawHeader(true);
  let y = H - 42 * MM;
  [
    ["Nom et prénom :", `${etudiant.prenom || ""} ${etudiant.nom || ""}`],
    ["CIN :", etudiant.cin || "—"],
    ["N° Inscription :", etudiant.numInscription || "—"],
    ["Filière :", etudiant.filiere || "—"],
    ["Promotion :", etudiant.promotion || "—"],
  ].forEach(([label, val]) => {
    doc.text(marge, y, label, { size: 10.5, color: [0.42, 0.42, 0.42] });
    doc.text(marge + 45 * MM, y, String(val), { size: 10.5, bold: true });
    y -= 6 * MM;
  });
  y -= 5 * MM;
  const total = presencesEtudiant.length;
  const justifiees = presencesEtudiant.filter((p) => p.statut === "Justifié").length;
  const nonJustifiees = total - justifiees;
  doc.rect(marge, y - 16 * MM, W - 2 * marge, 16 * MM, [0.95, 0.96, 0.98]);
  doc.text(marge + 5 * MM, y - 6.5 * MM, `Total absences : ${total}`, { size: 11, bold: true, color: NAVY_RGB });
  doc.text(marge + 70 * MM, y - 6.5 * MM, `Justifiées : ${justifiees}`, {
    size: 11,
    bold: true,
    color: [0.18, 0.49, 0.32],
  });
  doc.text(marge + 130 * MM, y - 6.5 * MM, `Non justifiées : ${nonJustifiees}`, {
    size: 11,
    bold: true,
    color: [0.75, 0.22, 0.17],
  });
  y -= 24 * MM;
  const cols = [
    { label: "Date", w: 24 * MM },
    { label: "Module", w: 58 * MM },
    { label: "Séance", w: 34 * MM },
    { label: "Statut", w: 24 * MM },
    { label: "Justification", w: W - 2 * marge - (24 + 58 + 34 + 24) * MM },
  ];
  const drawTableHeader = () => {
    doc.rect(marge, y - 7 * MM, W - 2 * marge, 7 * MM, NAVY_RGB);
    let cx = marge;
    cols.forEach((c) => {
      doc.text(cx + 2 * MM, y - 5 * MM, c.label, { size: 9, bold: true, color: [1, 1, 1] });
      cx += c.w;
    });
    y -= 7 * MM;
  };
  const drawSemestreHeader = (nom) => {
    doc.rect(marge, y - 6 * MM, W - 2 * marge, 6 * MM, [0.99, 0.93, 0.87]);
    doc.text(marge + 2 * MM, y - 4.3 * MM, nom, { size: 9.5, bold: true, color: ORANGE_RGB });
    y -= 6 * MM;
  };
  if (presencesEtudiant.length === 0) {
    drawTableHeader();
    doc.text(marge + 2 * MM, y - 4.3 * MM, "Aucune absence enregistrée.", { size: 10, color: [0.5, 0.5, 0.5] });
    y -= 6 * MM;
  } else {
    const parSemestre = {};
    presencesEtudiant.forEach((p) => {
      const moduleInfo = coursListe.find((c) => c.intitule === p.module);
      const key = moduleInfo?.semestre || p.semestre || "Semestre non précisé";
      if (!parSemestre[key]) parSemestre[key] = [];
      parSemestre[key].push(p);
    });
    const semestreKeys = Object.keys(parSemestre).sort((a, b) =>
      a.localeCompare(b, "fr", { numeric: true, sensitivity: "base" })
    );
    semestreKeys.forEach((sk) => {
      if (y < 40 * MM) {
        doc.newPage();
        drawHeader(false);
        y = H - 35 * MM;
      }
      drawSemestreHeader(sk);
      drawTableHeader();
      const trie = [...parSemestre[sk]].sort((a, b) => (b.date || "").localeCompare(a.date || ""));
      trie.forEach((p, i) => {
        if (y < 32 * MM) {
          doc.newPage();
          drawHeader(false);
          y = H - 35 * MM;
          drawSemestreHeader(`${sk} (suite)`);
          drawTableHeader();
        }
        if (i % 2 === 1) doc.rect(marge, y - 6 * MM, W - 2 * marge, 6 * MM, [0.96, 0.98, 0.97]);
        const vals = [
          formatDateFR(p.date) || "—",
          p.module || "—",
          p.seance || p.heure || "—",
          p.statut || "Absent",
          p.statut === "Justifié" ? p.documentJustification || "Justifié" : "—",
        ];
        let cx = marge;
        vals.forEach((val, ci) => {
          doc.text(cx + 2 * MM, y - 4.3 * MM, String(val).slice(0, ci === 1 ? 38 : 22), { size: 8.3 });
          cx += cols[ci].w;
        });
        y -= 6 * MM;
      });
      y -= 4 * MM;
    });
  }
  y -= 12 * MM;
  if (y < 40 * MM) {
    doc.newPage();
    drawHeader(false);
    y = H - 45 * MM;
  }
  doc.text(marge, y, `Fait à Rabat, le ${/* @__PURE__ */ new Date().toLocaleDateString("fr-FR")}`, { size: 10 });
  doc.text(W - marge - 58 * MM, y, "Le Responsable Pédagogique", { size: 11, bold: true, color: NAVY_RGB });
  doc.textCentered(W / 2, 8 * MM, "ESRMI — École Supérieure de Rabat en Management et Ingénierie", {
    size: 8,
    color: [0.5, 0.5, 0.5],
  });
  downloadPDF(doc.build(), `Releve_Absences_${etudiant.prenom || ""}_${etudiant.nom || ""}.pdf`.replace(/\s+/g, "_"));
}
function calculerLignesBulletin(
  etudiant,
  notesEtudiantSemestre,
  semestre,
  coursListe,
  presencesEtudiantSemestre,
  filieresAvecTauxAbsence = []
) {
  const modulesDuSemestre = (coursListe || []).filter(
    (c) => c.semestre === semestre && (c.filieresListe || [c.filiere]).includes(etudiant.filiere)
  );
  const tauxAbsencePriseEnCompte = filieresAvecTauxAbsence.includes(etudiant.filiere);
  let sommeNotesModules = 0,
    nbModulesNotes = 0;
  const brut = modulesDuSemestre.map((module) => {
    const notesModule = notesEtudiantSemestre.filter((n) => n.module === module.intitule);
    const ccNotes = notesModule.filter((n) => n.typeEvaluation === "Contrôle continu").map((n) => n.note);
    const efmNote = notesModule.find((n) => n.typeEvaluation === "Examen de fin de module")?.note ?? null;
    const rattrapageNote = notesModule.find((n) => n.typeEvaluation === "Rattrapage")?.note ?? null;
    const moyenneCC = ccNotes.length ? ccNotes.reduce((s, n) => s + n, 0) / ccNotes.length : null;
    const noteModuleInitiale = moyenneCC !== null && efmNote !== null ? (moyenneCC + efmNote) / 2 : null;
    const noteModule = rattrapageNote !== null ? rattrapageNote : noteModuleInitiale;
    let tauxAbsence = null;
    if (tauxAbsencePriseEnCompte) {
      const heuresManquees = (presencesEtudiantSemestre || [])
        .filter((p) => p.module === module.intitule)
        .reduce((s, p) => s + (SEANCE_DUREE[p.seance || p.heure] || 1.5), 0);
      tauxAbsence = module.heures > 0 ? Math.round((heuresManquees / module.heures) * 10000) / 100 : null;
    }
    if (noteModule !== null) {
      sommeNotesModules += noteModule;
      nbModulesNotes++;
    }
    return { module: module.intitule, moyenneCC, efmNote, noteModule, tauxAbsence, rattrapageNote };
  });
  const moyenneSemestre = nbModulesNotes > 0 ? sommeNotesModules / nbModulesNotes : null; // R\xE8gle de d\xE9cision du jury :
  // - Moyenne du semestre < 10 : seuls les modules avec une note < 10 sont non valid\xE9s
  //   (les modules \xE0 10 ou plus restent valid\xE9s, pas de compensation en dessous de la moyenne)
  // - Moyenne du semestre >= 10 : compensation activ\xE9e, un module avec une note >= 6 est valid\xE9
  const lignes = brut.map((l) => {
    let decision = "";
    if (l.noteModule !== null) {
      if (moyenneSemestre !== null && moyenneSemestre < 10) {
        decision = l.noteModule < 10 ? "Non validé" : "Validé";
      } else {
        decision = l.noteModule >= 6 ? "Validé" : "Non validé";
      }
    }
    if (l.tauxAbsence !== null && l.tauxAbsence >= 25) decision = "Non validé";
    return { ...l, decision };
  });
  return { lignes, moyenneSemestre, tauxAbsencePriseEnCompte };
}
function genererBulletinNotes(
  etudiant,
  notesEtudiantSemestre,
  semestre,
  anneeUniversitaire,
  coursListe,
  presencesEtudiantSemestre,
  filieresAvecTauxAbsence = [],
  docExterne = null
) {
  const W = 210 * MM,
    H = 297 * MM;
  const marge = 18 * MM;
  const doc = docExterne || new MiniPDF(W, H);
  const drawBox = (x, y2, w, h, color = [0.55, 0.55, 0.55]) => {
    doc.line(x, y2, x + w, y2, 0.7, color);
    doc.line(x + w, y2, x + w, y2 + h, 0.7, color);
    doc.line(x + w, y2 + h, x, y2 + h, 0.7, color);
    doc.line(x, y2 + h, x, y2, 0.7, color);
  };
  const fmt = (v) => (v === null || v === void 0 || isNaN(v) ? "" : v.toFixed(2).replace(".", ","));
  const truncateFit = (text, maxW, size, bold = false) => {
    let t = String(text ?? "—");
    if (measureTextWidth(t, size, bold) <= maxW) return t;
    while (t.length > 1 && measureTextWidth(t + "…", size, bold) > maxW) t = t.slice(0, -1);
    return t + "…";
  };
  const logoW = 20 * MM,
    logoH = 16.875 * MM;
  const headerBottomY = H - 30 * MM;
  doc.image(base64ToBytes(LOGO_ESR_B64), marge, headerBottomY, logoW, logoH);
  doc.textCentered(W / 2, headerBottomY + logoH / 2 - 2 * MM, "Bulletin de Notes", {
    size: 21,
    bold: true,
    color: [0, 0, 0],
  });
  doc.line(marge, headerBottomY - 4 * MM, W - marge, headerBottomY - 4 * MM, 1.3, [0, 0, 0]);
  let y = headerBottomY - 15 * MM;
  doc.textCentered(W / 2, y, `${etudiant.prenom || ""} ${etudiant.nom || ""}`.toUpperCase(), {
    size: 17,
    bold: true,
    color: NAVY_RGB,
  });
  y -= 11 * MM;
  const partiesFiliere = (etudiant.filiere || "")
    .split(",")
    .map((s) => s.trim())
    .filter(Boolean);
  if (partiesFiliere.length > 0) {
    doc.textCentered(W / 2, y, partiesFiliere[0], { size: 12, bold: true });
    y -= 7 * MM;
  }
  if (partiesFiliere.length > 1) {
    doc.textCentered(W / 2, y, partiesFiliere.slice(1).join(", "), { size: 11, bold: true });
    y -= 7 * MM;
  }
  const niveauAffichable = etudiant.niveau && !/^BAC\s*\+/i.test(etudiant.niveau.trim()) ? etudiant.niveau : "";
  doc.textCentered(
    W / 2,
    y,
    `${niveauAffichable ? niveauAffichable + "  -  " : ""}${semestre}      Année universitaire: ${anneeUniversitaire}`,
    { size: 10.5, bold: true }
  );
  y -= 13 * MM;
  const { lignes, moyenneSemestre, tauxAbsencePriseEnCompte } = calculerLignesBulletin(
    etudiant,
    notesEtudiantSemestre,
    semestre,
    coursListe,
    presencesEtudiantSemestre,
    filieresAvecTauxAbsence
  );
  const cols = tauxAbsencePriseEnCompte
    ? [
        { label: "Module", w: 48 * MM },
        { label: "Moyenne CC/TP", w: 20 * MM },
        { label: "Note EFM", w: 17 * MM },
        { label: "Note Module", w: 17 * MM },
        { label: "Taux d'absence atteint", w: 22 * MM },
        { label: "Note rattrapage", w: 20 * MM },
        { label: "Décision Jury d'examen", w: W - 2 * marge - (48 + 20 + 17 + 17 + 22 + 20) * MM },
      ]
    : [
        { label: "Module", w: 58 * MM },
        { label: "Moyenne CC/TP", w: 22 * MM },
        { label: "Note EFM", w: 19 * MM },
        { label: "Note Module", w: 19 * MM },
        { label: "Note rattrapage", w: 22 * MM },
        { label: "Décision Jury d'examen", w: W - 2 * marge - (58 + 22 + 19 + 19 + 22) * MM },
      ];
  const tableX = marge,
    tableW = W - 2 * marge,
    headerH = 13 * MM,
    rowH = 8 * MM;
  const RED_RGB = [0.85, 0.16, 0.16];
  const motifPour = (l) =>
    l.tauxAbsence !== null && l.tauxAbsence >= 25 ? "Seuil d'absences dépassé" : "Note insuffisante";
  const drawTableHeader = () => {
    doc.rect(tableX, y - headerH, tableW, headerH, [0, 0, 0]);
    let cx = tableX;
    cols.forEach((c) => {
      const lignesLabel = wrapText(c.label, c.w - 2 * MM, 7, true);
      const ligneH = 7 * 1.3;
      const startY = y - headerH / 2 + ((lignesLabel.length - 1) * ligneH) / 2 - 1.5;
      lignesLabel.forEach((ligne, li) => {
        doc.textCentered(cx + c.w / 2, startY - li * ligneH, ligne, { size: 7, bold: true, color: [1, 1, 1] });
      });
      cx += c.w;
    });
    y -= headerH;
  };
  let tableTopY = y;
  let rowBottoms = [];
  const dessinerGrilleTableau = (hautPage) => {
    doc.line(tableX, hautPage, tableX, y, 0.9, [0, 0, 0]);
    doc.line(tableX + tableW, hautPage, tableX + tableW, y, 0.9, [0, 0, 0]);
    doc.line(tableX, hautPage, tableX + tableW, hautPage, 0.9, [0, 0, 0]);
    rowBottoms.forEach((ry) => doc.line(tableX, ry, tableX + tableW, ry, 0.4, [0.55, 0.55, 0.55]));
    let cx2 = tableX;
    doc.line(cx2, hautPage, cx2, y, 0.4, [0.55, 0.55, 0.55]);
    cols.forEach((c) => {
      cx2 += c.w;
      doc.line(cx2, hautPage, cx2, y, 0.4, [0.55, 0.55, 0.55]);
    });
    doc.line(tableX, y, tableX + tableW, y, 0.9, [0, 0, 0]);
  };
  drawTableHeader();
  lignes.forEach((l, i) => {
    if (y < 45 * MM) {
      dessinerGrilleTableau(tableTopY);
      doc.newPage();
      doc.rect(0, H - 1.2 * MM, W, 1.2 * MM, ORANGE_RGB);
      y = H - 30 * MM;
      tableTopY = y;
      drawTableHeader();
      rowBottoms = [];
    }
    if (i % 2 === 1) doc.rect(tableX, y - rowH, tableW, rowH, [0.96, 0.96, 0.97]);
    const vals = tauxAbsencePriseEnCompte
      ? [
          l.module,
          fmt(l.moyenneCC),
          fmt(l.efmNote),
          fmt(l.noteModule),
          l.tauxAbsence !== null ? `${l.tauxAbsence.toFixed(2)}%` : "—",
          fmt(l.rattrapageNote),
          l.decision,
        ]
      : [l.module, fmt(l.moyenneCC), fmt(l.efmNote), fmt(l.noteModule), fmt(l.rattrapageNote), l.decision];
    let cx = tableX;
    vals.forEach((v, ci) => {
      const maxW = cols[ci].w - 3 * MM;
      const texte = truncateFit(v, maxW, 8);
      const estColDecision = ci === cols.length - 1;
      const estColTaux = tauxAbsencePriseEnCompte && cols[ci].label === "Taux d'absence atteint";
      const enRouge =
        l.decision === "Non validé" &&
        (estColDecision || (estColTaux && l.tauxAbsence !== null && l.tauxAbsence >= 25));
      const opts = { size: 8, color: enRouge ? RED_RGB : [0, 0, 0], bold: enRouge };
      if (ci === 0) doc.text(cx + 1.5 * MM, y - rowH / 2 - 1.3, texte, opts);
      else doc.textCentered(cx + cols[ci].w / 2, y - rowH / 2 - 1.3, texte, opts);
      cx += cols[ci].w;
    });
    y -= rowH;
    rowBottoms.push(y);
  });
  dessinerGrilleTableau(tableTopY);
  y -= 11 * MM;
  if (y < 55 * MM) {
    doc.newPage();
    doc.rect(0, H - 1.2 * MM, W, 1.2 * MM, ORANGE_RGB);
    y = H - 40 * MM;
  }
  const labelMoyenne = "Moyenne du semestre :";
  const valeurMoyenne = moyenneSemestre !== null ? fmt(moyenneSemestre) : "—";
  const labelW = measureTextWidth(labelMoyenne, 12.5, true);
  const valeurW = measureTextWidth(valeurMoyenne, 13.5, true);
  const gapMoyenne = 6 * MM;
  const startXMoyenne = W / 2 - (labelW + gapMoyenne + valeurW) / 2;
  doc.text(startXMoyenne, y, labelMoyenne, { size: 12.5, bold: true });
  doc.text(startXMoyenne + labelW + gapMoyenne, y, valeurMoyenne, { size: 13.5, bold: true, color: NAVY_RGB });
  y -= 14 * MM;
  const modulesAjournes = lignes.filter((l) => l.decision === "Non validé");
  if (modulesAjournes.length > 0) {
    if (y < 55 * MM) {
      doc.newPage();
      doc.rect(0, H - 1.2 * MM, W, 1.2 * MM, ORANGE_RGB);
      y = H - 40 * MM;
    }
    const colsAj = [
      { label: "Module", w: tableW * 0.62 },
      { label: "Motif", w: tableW - tableW * 0.62 },
    ];
    const ajHeaderH = 8 * MM,
      ajRowH = 7.5 * MM;
    doc.rect(tableX, y - ajHeaderH, tableW, ajHeaderH, RED_RGB);
    doc.text(tableX + 2 * MM, y - ajHeaderH / 2 - 1.5, "Liste des modules ajournés — à repasser", {
      size: 9,
      bold: true,
      color: [1, 1, 1],
    });
    doc.textCentered(tableX + colsAj[0].w + colsAj[1].w / 2, y - ajHeaderH / 2 - 1.5, "Motif", {
      size: 9,
      bold: true,
      color: [1, 1, 1],
    });
    y -= ajHeaderH;
    let ajTopY = y + ajHeaderH;
    let ajRowBottoms = [];
    const dessinerGrilleAj = (hautPage) => {
      doc.line(tableX, hautPage, tableX, y, 0.9, [0, 0, 0]);
      doc.line(tableX + tableW, hautPage, tableX + tableW, y, 0.9, [0, 0, 0]);
      doc.line(tableX, hautPage, tableX + tableW, hautPage, 0.9, [0, 0, 0]);
      ajRowBottoms.forEach((ry) => doc.line(tableX, ry, tableX + tableW, ry, 0.4, [0.55, 0.55, 0.55]));
      doc.line(tableX, hautPage, tableX, y, 0.4, [0.55, 0.55, 0.55]);
      doc.line(tableX + colsAj[0].w, hautPage, tableX + colsAj[0].w, y, 0.4, [0.55, 0.55, 0.55]);
      doc.line(tableX + tableW, hautPage, tableX + tableW, y, 0.4, [0.55, 0.55, 0.55]);
      doc.line(tableX, y, tableX + tableW, y, 0.9, [0, 0, 0]);
    };
    modulesAjournes.forEach((l, i) => {
      if (y < 30 * MM) {
        dessinerGrilleAj(ajTopY);
        doc.newPage();
        doc.rect(0, H - 1.2 * MM, W, 1.2 * MM, ORANGE_RGB);
        y = H - 30 * MM;
        ajTopY = y;
        ajRowBottoms = [];
      }
      if (i % 2 === 1) doc.rect(tableX, y - ajRowH, tableW, ajRowH, [0.96, 0.96, 0.97]);
      doc.text(tableX + 2 * MM, y - ajRowH / 2 - 1.3, truncateFit(l.module, colsAj[0].w - 4 * MM, 9), { size: 9 });
      doc.textCentered(tableX + colsAj[0].w + colsAj[1].w / 2, y - ajRowH / 2 - 1.3, motifPour(l), { size: 9 });
      y -= ajRowH;
      ajRowBottoms.push(y);
    });
    dessinerGrilleAj(ajTopY);
    y -= 23 * MM;
  } else {
    y -= 17 * MM;
  }
  if (y < 25 * MM) y = 25 * MM;
  const boxW = 75 * MM,
    boxH = 15 * MM;
  drawBox(marge, y, boxW, boxH);
  doc.textCentered(marge + boxW / 2, y + boxH / 2 - 1.5 * MM, "Signature du directeur pédagogique", {
    size: 8.5,
    bold: true,
  });
  drawBox(W - marge - boxW, y, boxW, boxH);
  doc.textCentered(W - marge - boxW / 2, y + boxH / 2 - 1.5 * MM, "Signature du président de jury d'examen", {
    size: 8.5,
    bold: true,
  });
  if (docExterne) return;
  downloadPDF(
    doc.build(),
    `Bulletin_${etudiant.prenom || ""}_${etudiant.nom || ""}_${semestre}`.replace(/\s+/g, "_") + ".pdf"
  );
}
function genererReleveDetaille(etudiant, notesEtudiantSemestre, semestre, anneeUniversitaire, docExterne = null) {
  const W = 210 * MM,
    H = 297 * MM;
  const marge = 18 * MM;
  const doc = docExterne || new MiniPDF(W, H);
  const fmt = (v) => (v === null || v === void 0 || isNaN(v) ? "—" : v.toFixed(2).replace(".", ","));
  const truncateFit = (text, maxW, size, bold = false) => {
    let t = String(text ?? "—");
    if (measureTextWidth(t, size, bold) <= maxW) return t;
    while (t.length > 1 && measureTextWidth(t + "…", size, bold) > maxW) t = t.slice(0, -1);
    return t + "…";
  };
  const logoW = 20 * MM,
    logoH = 16.875 * MM;
  const drawHeader = (premierePage) => {
    const headerBottomY = H - 30 * MM;
    if (premierePage) {
      doc.image(base64ToBytes(LOGO_ESR_B64), marge, headerBottomY, logoW, logoH);
      doc.textCentered(W / 2, headerBottomY + logoH / 2 - 2 * MM, "Relevé Détaillé", {
        size: 21,
        bold: true,
        color: [0, 0, 0],
      });
    } else {
      doc.textCentered(W / 2, headerBottomY + logoH / 2 - 2 * MM, "Relevé Détaillé des évaluations (suite)", {
        size: 14,
        bold: true,
        color: [0, 0, 0],
      });
    }
    doc.line(marge, headerBottomY - 4 * MM, W - marge, headerBottomY - 4 * MM, 1.3, [0, 0, 0]);
  };
  drawHeader(true);
  let y = H - 53 * MM;
  doc.textCentered(W / 2, y, `${etudiant.prenom || ""} ${etudiant.nom || ""}`.toUpperCase(), {
    size: 15,
    bold: true,
    color: NAVY_RGB,
  });
  y -= 9 * MM;
  const anneeLabel = `Année universitaire : ${anneeUniversitaire}`;
  const anneeW = measureTextWidth(anneeLabel, 10.5, true);
  const filiereLigne = `${etudiant.filiere || ""} — ${semestre}`;
  const maxFiliereW = W - 2 * marge - anneeW - 8 * MM;
  doc.text(marge, y, truncateFit(filiereLigne, maxFiliereW, 10.5, true), { size: 10.5, bold: true });
  doc.textRight(W - marge, y, anneeLabel, { size: 10.5, bold: true });
  y -= 12 * MM;
  const tableX = marge,
    tableW = W - 2 * marge;
  const cols = [
    { label: "Date", w: 24 * MM },
    { label: "Module", w: 74 * MM },
    { label: "Type d'évaluation", w: 42 * MM },
    { label: "Note /20", w: tableW - (24 + 74 + 42) * MM },
  ];
  const headerH = 9 * MM,
    rowH = 7.5 * MM;
  const drawTableHeader = () => {
    doc.rect(tableX, y - headerH, tableW, headerH, NAVY_RGB);
    let cx = tableX;
    cols.forEach((c) => {
      doc.textCentered(cx + c.w / 2, y - headerH / 2 - 2, truncateFit(c.label, c.w - 2 * MM, 8, true), {
        size: 8,
        bold: true,
        color: [1, 1, 1],
      });
      cx += c.w;
    });
    y -= headerH;
  };
  let tableTopY = y;
  let rowBottoms = [];
  const dessinerGrilleTableau = (hautPage) => {
    doc.line(tableX, hautPage, tableX, y, 0.9, [0, 0, 0]);
    doc.line(tableX + tableW, hautPage, tableX + tableW, y, 0.9, [0, 0, 0]);
    doc.line(tableX, hautPage, tableX + tableW, hautPage, 0.9, [0, 0, 0]);
    rowBottoms.forEach((ry) => doc.line(tableX, ry, tableX + tableW, ry, 0.4, [0.6, 0.6, 0.6]));
    let cx2 = tableX;
    doc.line(cx2, hautPage, cx2, y, 0.4, [0.6, 0.6, 0.6]);
    cols.forEach((c) => {
      cx2 += c.w;
      doc.line(cx2, hautPage, cx2, y, 0.4, [0.6, 0.6, 0.6]);
    });
    doc.line(tableX, y, tableX + tableW, y, 0.9, [0, 0, 0]);
  };
  drawTableHeader();
  const trie = [...notesEtudiantSemestre].sort(
    (a, b) =>
      (b.date || "").localeCompare(a.date || "") ||
      (a.module || "").localeCompare(b.module || "", "fr", { sensitivity: "base" })
  );
  if (trie.length === 0) {
    doc.text(tableX + 2 * MM, y - 5 * MM, "Aucune évaluation enregistrée.", { size: 10, color: [0.5, 0.5, 0.5] });
    y -= 7 * MM;
  }
  trie.forEach((n, i) => {
    if (y < 35 * MM) {
      dessinerGrilleTableau(tableTopY);
      doc.newPage();
      drawHeader(false);
      y = H - 45 * MM;
      tableTopY = y;
      drawTableHeader();
      rowBottoms = [];
    }
    if (i % 2 === 1) doc.rect(tableX, y - rowH, tableW, rowH, [0.96, 0.96, 0.98]);
    const vals = [formatDateFR(n.date) || "—", n.module || "—", n.typeEvaluation || "—", fmt(n.note)];
    let cx = tableX;
    vals.forEach((v, ci) => {
      const maxW = cols[ci].w - 3 * MM;
      const texte = truncateFit(v, maxW, 8.5);
      if (ci === 1) doc.text(cx + 1.8 * MM, y - rowH / 2 - 1.2, texte, { size: 8.5 });
      else doc.textCentered(cx + cols[ci].w / 2, y - rowH / 2 - 1.2, texte, { size: 8.5 });
      cx += cols[ci].w;
    });
    y -= rowH;
    rowBottoms.push(y);
  });
  dessinerGrilleTableau(tableTopY);
  doc.textCentered(W / 2, 8 * MM, "ESRMI — École Supérieure de Rabat en Management et Ingénierie", {
    size: 8,
    color: [0.5, 0.5, 0.5],
  });
  if (docExterne) return;
  downloadPDF(
    doc.build(),
    `Releve_Detaille_${etudiant.prenom || ""}_${etudiant.nom || ""}_${semestre}`.replace(/\s+/g, "_") + ".pdf"
  );
}
function genererFichePointage(fiche, seances, professeur, formatMoisLabelFn, periodeFn) {
  const W = 210 * MM,
    H = 297 * MM;
  const marge = 18 * MM;
  const doc = new MiniPDF(W, H);
  const fmt2 = (v) =>
    v === null || v === void 0 || isNaN(v) ? "—" : Number(v).toFixed(2).replace(/\.00$/, "").replace(".", ",");
  const truncateFit = (text, maxW, size, bold = false) => {
    let t = String(text ?? "—");
    if (measureTextWidth(t, size, bold) <= maxW) return t;
    while (t.length > 1 && measureTextWidth(t + "…", size, bold) > maxW) t = t.slice(0, -1);
    return t + "…";
  };
  const logoW = 20 * MM,
    logoH = 16.875 * MM;
  const headerBottomY = H - 30 * MM;
  doc.image(base64ToBytes(LOGO_ESR_B64), marge, headerBottomY, logoW, logoH);
  doc.textCentered(W / 2, headerBottomY + logoH / 2 - 2 * MM, "Fiche de Pointage", {
    size: 21,
    bold: true,
    color: [0, 0, 0],
  });
  doc.line(marge, headerBottomY - 4 * MM, W - marge, headerBottomY - 4 * MM, 1.3, [0, 0, 0]);
  let y = H - 53 * MM;
  doc.textCentered(
    W / 2,
    y,
    `${(professeur?.prenom || fiche.professeurNom || "").toUpperCase()} ${(professeur?.nom || "").toUpperCase()}`.trim(),
    { size: 15, bold: true, color: NAVY_RGB }
  );
  y -= 9 * MM;
  const periode = periodeFn ? periodeFn(fiche.mois) : { debut: "", fin: "" };
  doc.textCentered(
    W / 2,
    y,
    `${formatMoisLabelFn ? formatMoisLabelFn(fiche.mois) : fiche.mois}  —  Période du ${formatDateFR(periode.debut) || "—"} au ${formatDateFR(periode.fin) || "—"}`,
    { size: 10.5, bold: true }
  );
  y -= 13 * MM;
  const tableX = marge,
    tableW = W - 2 * marge;
  const cols = [
    { label: "Date", w: 22 * MM },
    { label: "Module", w: 62 * MM },
    { label: "Horaires", w: 32 * MM },
    { label: "Durée", w: 18 * MM },
    { label: "Commentaire", w: tableW - (22 + 62 + 32 + 18) * MM },
  ];
  const headerH = 9 * MM,
    rowH = 8 * MM;
  const drawTableHeader = () => {
    doc.rect(tableX, y - headerH, tableW, headerH, [0, 0, 0]);
    let cx = tableX;
    cols.forEach((c) => {
      doc.textCentered(cx + c.w / 2, y - headerH / 2 - 2, c.label, { size: 8, bold: true, color: [1, 1, 1] });
      cx += c.w;
    });
    y -= headerH;
  };
  let tableTopY = y;
  let rowBottoms = [];
  const dessinerGrilleTableauFiche = (hautPage) => {
    doc.line(tableX, hautPage, tableX, y, 0.9, [0, 0, 0]);
    doc.line(tableX + tableW, hautPage, tableX + tableW, y, 0.9, [0, 0, 0]);
    doc.line(tableX, hautPage, tableX + tableW, hautPage, 0.9, [0, 0, 0]);
    rowBottoms.forEach((ry) => doc.line(tableX, ry, tableX + tableW, ry, 0.4, [0.55, 0.55, 0.55]));
    let cx2 = tableX;
    doc.line(cx2, hautPage, cx2, y, 0.4, [0.55, 0.55, 0.55]);
    cols.forEach((c) => {
      cx2 += c.w;
      doc.line(cx2, hautPage, cx2, y, 0.4, [0.55, 0.55, 0.55]);
    });
    doc.line(tableX, y, tableX + tableW, y, 0.9, [0, 0, 0]);
  };
  drawTableHeader();
  const seancesTriees = [...(seances || [])].sort((a, b) => (a.date || "").localeCompare(b.date || ""));
  if (seancesTriees.length === 0) {
    doc.text(tableX + 2 * MM, y - 5 * MM, "Aucune séance enregistrée.", { size: 10, color: [0.5, 0.5, 0.5] });
    y -= 7 * MM;
  }
  seancesTriees.forEach((s, i) => {
    if (y < 45 * MM) {
      dessinerGrilleTableauFiche(tableTopY);
      doc.newPage();
      doc.rect(0, H - 1.2 * MM, W, 1.2 * MM, ORANGE_RGB);
      y = H - 30 * MM;
      tableTopY = y;
      drawTableHeader();
      rowBottoms = [];
    }
    if (i % 2 === 1) doc.rect(tableX, y - rowH, tableW, rowH, [0.96, 0.96, 0.97]);
    const vals = [
      formatDateFR(s.date) || "—",
      s.module || "—",
      `${s.heureDebut || ""}-${s.heureFin || ""}`,
      `${fmt2(s.nbHeures)}h`,
      s.commentaire || "",
    ];
    let cx = tableX;
    vals.forEach((v, ci) => {
      const maxW = cols[ci].w - 3 * MM;
      const texte = truncateFit(v, maxW, 8);
      if (ci === 1 || ci === 4) doc.text(cx + 1.5 * MM, y - rowH / 2 - 1.3, texte, { size: 8 });
      else doc.textCentered(cx + cols[ci].w / 2, y - rowH / 2 - 1.3, texte, { size: 8 });
      cx += cols[ci].w;
    });
    y -= rowH;
    rowBottoms.push(y);
  });
  dessinerGrilleTableauFiche(tableTopY);
  y -= 11 * MM;
  if (y < 70 * MM) {
    doc.newPage();
    doc.rect(0, H - 1.2 * MM, W, 1.2 * MM, ORANGE_RGB);
    y = H - 40 * MM;
  }
  const totalHeures = (seances || []).reduce((s, x) => s + (x.nbHeures || 0), 0);
  doc.text(tableX, y, "Total des heures :", { size: 12.5, bold: true });
  doc.text(tableX + 55 * MM, y, `${fmt2(totalHeures)}h`, { size: 13.5, bold: true, color: NAVY_RGB });
  y -= 14 * MM;
  doc.text(tableX, y, "Circuit de validation", { size: 11, bold: true, color: NAVY_RGB });
  y -= 8 * MM;
  const etapes = [
    {
      label: "Soumission (professeur)",
      fait: !!fiche.dateSoumission,
      qui: fiche.professeurNom,
      quand: fiche.dateSoumission,
    },
    {
      label: "Validation des séances (Iliass Elmzioui)",
      fait: !!fiche.dateValidation,
      qui: fiche.valideePar,
      quand: fiche.dateValidation,
    },
    {
      label: "Accusé de réception (Kawtar Bouabid)",
      fait: !!fiche.dateAccuseReception,
      qui: fiche.accuseReceptionPar,
      quand: fiche.dateAccuseReception,
    },
    {
      label: "Traitement du paiement (Fatima Zahra Zniber)",
      fait: !!fiche.dateTraitement,
      qui: fiche.traitePar,
      quand: fiche.dateTraitement,
    },
  ];
  etapes.forEach((e) => {
    const puce = e.fait ? [0.16, 0.6, 0.32] : [0.75, 0.75, 0.75];
    doc.rect(tableX, y - 3.2 * MM, 3.2 * MM, 3.2 * MM, puce);
    doc.text(tableX + 6 * MM, y - 2.4 * MM, e.label, {
      size: 9.5,
      bold: e.fait,
      color: e.fait ? NAVY_RGB : [0.5, 0.5, 0.5],
    });
    if (e.fait)
      doc.text(tableX + 6 * MM, y - 6.2 * MM, `— ${e.qui || ""} · ${formatDateFR(e.quand) || ""}`, {
        size: 8,
        color: [0.5, 0.5, 0.5],
      });
    y -= e.fait ? 12 * MM : 8 * MM;
  });
  doc.textCentered(W / 2, 8 * MM, "ESRMI — École Supérieure de Rabat en Management et Ingénierie", {
    size: 8,
    color: [0.5, 0.5, 0.5],
  });
  downloadPDF(
    doc.build(),
    `Pointage_${(professeur?.nom || fiche.professeurNom || "").replace(/\s+/g, "_")}_${fiche.mois}.pdf`
  );
}
function genererFichePresenceHebdo(filiereObj, etudiantsListe, lundiStr, JOURS_SEMAINE_FR) {
  const W = 297 * MM,
    H = 210 * MM;
  const marge = 14 * MM;
  const doc = new MiniPDF(W, H);
  const truncateFit = (text, maxW, size, bold = false) => {
    let t = String(text ?? "—");
    if (measureTextWidth(t, size, bold) <= maxW) return t;
    while (t.length > 1 && measureTextWidth(t + "…", size, bold) > maxW) t = t.slice(0, -1);
    return t + "…";
  };
  const ajouterJours = (dateStr, n) => {
    const [an, mois, jourNum] = dateStr.split("-").map(Number);
    const d = new Date(an, mois - 1, jourNum);
    d.setDate(d.getDate() + n);
    return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
  };
  const vendrediStr = ajouterJours(lundiStr, 4);
  const logoW = 16 * MM,
    logoH = 13.5 * MM;
  const headerBottomY = H - 24 * MM;
  doc.image(base64ToBytes(LOGO_ESR_B64), marge, headerBottomY, logoW, logoH);
  doc.textCentered(W / 2, headerBottomY + logoH / 2 - 2 * MM, "Fiche de Présence Hebdomadaire", {
    size: 18,
    bold: true,
    color: [0, 0, 0],
  });
  doc.line(marge, headerBottomY - 4 * MM, W - marge, headerBottomY - 4 * MM, 1.2, [0, 0, 0]);
  let y = H - 36 * MM;
  doc.text(marge, y, `Groupe : ${filiereObj?.intitule || "—"}${filiereObj?.code ? " (" + filiereObj.code + ")" : ""}`, {
    size: 11,
    bold: true,
    color: NAVY_RGB,
  });
  doc.textRight(W - marge, y, `Semaine du ${formatDateFR(lundiStr)} au ${formatDateFR(vendrediStr)}`, {
    size: 11,
    bold: true,
    color: NAVY_RGB,
  });
  y -= 10 * MM;
  const slots =
    filiereObj?.typeSeances === "Jour"
      ? ["09h00 - 10h30", "10h45 - 12h15", "14h00 - 15h30", "15h45 - 17h15"]
      : ["19h00 - 21h00"];
  const nbSlots = slots.length;
  const jours = JOURS_SEMAINE_FR || ["Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi"];
  const colNum = 9 * MM,
    colNom = 62 * MM;
  const tableX = marge,
    tableW = W - 2 * marge;
  const largeurJours = tableW - colNum - colNom;
  const largeurJour = largeurJours / 5;
  const largeurSlot = largeurJour / nbSlots;
  const hauteurJourLabel = 6 * MM;
  const headerH = nbSlots > 1 ? 26 * MM : 14 * MM;
  const rowH = 8.5 * MM;
  const texteVertical = (centreX, yBas, yHaut, str, size, bold) => {
    const largeurTexte = measureTextWidth(str, size, bold);
    const bandMid = (yBas + yHaut) / 2;
    const yStart = bandMid - largeurTexte / 2;
    doc.text(centreX + size * 0.35, yStart, str, { size, bold, color: [1, 1, 1], angle: 90 });
  };
  const GRIS_ENTETE = [0.4, 0.4, 0.42];
  const drawTableHeader = () => {
    doc.rect(tableX, y - headerH, tableW, headerH, GRIS_ENTETE);
    doc.textCentered(tableX + colNum / 2, y - headerH / 2 - 2, "N°", { size: 8, bold: true, color: [1, 1, 1] });
    doc.textCentered(tableX + colNum + colNom / 2, y - headerH / 2 - 2, "Étudiant", {
      size: 8,
      bold: true,
      color: [1, 1, 1],
    });
    let cx = tableX + colNum + colNom;
    jours.forEach((jour, ji) => {
      const dateJour = ajouterJours(lundiStr, ji);
      doc.textCentered(cx + largeurJour / 2, y - hauteurJourLabel / 2 - 2, `${jour} ${formatDateFR(dateJour)}`, {
        size: 7.5,
        bold: true,
        color: [1, 1, 1],
      });
      let sx = cx;
      slots.forEach((s) => {
        if (nbSlots > 1) {
          texteVertical(sx + largeurSlot / 2, y - headerH + 1 * MM, y - hauteurJourLabel - 1 * MM, s, 6.5, true);
        } else {
          doc.textCentered(sx + largeurSlot / 2, y - headerH + (headerH - hauteurJourLabel) / 2 - 2, s, {
            size: 8,
            bold: true,
            color: [1, 1, 1],
          });
        }
        sx += largeurSlot;
      });
      cx += largeurJour;
    });
    y -= headerH;
  };
  let tableTopY = y;
  let rowBottoms = [];
  const dessinerGrille = (hautPage) => {
    doc.line(tableX, hautPage, tableX, y, 1, [0, 0, 0]);
    doc.line(tableX + tableW, hautPage, tableX + tableW, y, 1, [0, 0, 0]);
    doc.line(tableX, hautPage, tableX + tableW, hautPage, 1, [0, 0, 0]);
    doc.line(tableX, y, tableX + tableW, y, 1, [0, 0, 0]);
    rowBottoms.forEach((ry) => doc.line(tableX, ry, tableX + tableW, ry, 0.35, [0.6, 0.6, 0.6]));
    doc.line(tableX + colNum, hautPage, tableX + colNum, y, 0.6, [0, 0, 0]);
    doc.line(tableX + colNum + colNom, hautPage, tableX + colNum + colNom, y, 0.9, [0, 0, 0]);
    const basBandeJours = hautPage - hauteurJourLabel;
    let cx = tableX + colNum + colNom;
    for (let ji = 0; ji < 5; ji++) {
      for (let si = 0; si < nbSlots; si++) {
        cx += largeurSlot;
        const estFrontiereJour = si === nbSlots - 1;
        doc.line(
          cx,
          estFrontiereJour ? hautPage : basBandeJours,
          cx,
          y,
          estFrontiereJour ? 0.9 : 0.3,
          estFrontiereJour ? [0, 0, 0] : [0.7, 0.7, 0.7]
        );
      }
    }
  };
  drawTableHeader();
  const liste = [...(etudiantsListe || [])].sort((a, b) =>
    (a.nom || "").localeCompare(b.nom || "", "fr", { sensitivity: "base" })
  );
  if (liste.length === 0) {
    doc.text(tableX + 2 * MM, y - 5 * MM, "Aucun étudiant dans ce groupe.", { size: 10, color: [0.5, 0.5, 0.5] });
    y -= 7 * MM;
  }
  liste.forEach((et, i) => {
    if (y < 25 * MM) {
      dessinerGrille(tableTopY);
      doc.newPage();
      doc.rect(0, H - 1.2 * MM, W, 1.2 * MM, ORANGE_RGB);
      y = H - 20 * MM;
      tableTopY = y;
      drawTableHeader();
      rowBottoms = [];
    }
    if (i % 2 === 1) doc.rect(tableX, y - rowH, tableW, rowH, [0.96, 0.96, 0.97]);
    doc.textCentered(tableX + colNum / 2, y - rowH / 2 - 1.3, String(i + 1), { size: 8 });
    doc.text(
      tableX + colNum + 1.5 * MM,
      y - rowH / 2 - 1.3,
      truncateFit(`${(et.nom || "").toUpperCase()} ${et.prenom || ""}`, colNom - 3 * MM, 8),
      { size: 8 }
    );
    y -= rowH;
    rowBottoms.push(y);
  });
  dessinerGrille(tableTopY);
  doc.textCentered(
    W / 2,
    7 * MM,
    "ESRMI — École Supérieure de Rabat en Management et Ingénierie — À remplir manuellement (P = Présent, A = Absent)",
    { size: 7.5, color: [0.5, 0.5, 0.5] }
  );
  downloadPDF(
    doc.build(),
    `Fiche_Presence_${(filiereObj?.code || filiereObj?.intitule || "groupe").replace(/\s+/g, "_")}_${lundiStr}.pdf`
  );
}
function genererListeEtudiantsFiliere(filiereObj, etudiantsListe) {
  const W = 210 * MM,
    H = 297 * MM;
  const marge = 18 * MM;
  const doc = new MiniPDF(W, H);
  const truncateFit = (text, maxW, size, bold = false) => {
    let t = String(text ?? "—");
    if (measureTextWidth(t, size, bold) <= maxW) return t;
    while (t.length > 1 && measureTextWidth(t + "…", size, bold) > maxW) t = t.slice(0, -1);
    return t + "…";
  };
  const logoW = 20 * MM,
    logoH = 16.875 * MM;
  const headerBottomY = H - 30 * MM;
  doc.image(base64ToBytes(LOGO_ESR_B64), marge, headerBottomY, logoW, logoH);
  doc.textCentered(W / 2, headerBottomY + logoH / 2 - 2 * MM, "Liste des Étudiants", {
    size: 21,
    bold: true,
    color: [0, 0, 0],
  });
  doc.line(marge, headerBottomY - 4 * MM, W - marge, headerBottomY - 4 * MM, 1.3, [0, 0, 0]);
  let y = H - 53 * MM;
  const liste = [...(etudiantsListe || [])].sort((a, b) =>
    (a.nom || "").localeCompare(b.nom || "", "fr", { sensitivity: "base" })
  );
  doc.textCentered(W / 2, y, `${filiereObj?.intitule || "—"}${filiereObj?.code ? " (" + filiereObj.code + ")" : ""}`, {
    size: 13,
    bold: true,
    color: NAVY_RGB,
  });
  y -= 13 * MM;
  const tableX = marge,
    tableW = W - 2 * marge;
  const cols = [
    { label: "N°", w: 20 * MM },
    { label: "N° Inscription", w: 55 * MM },
    { label: "Nom et prénom", w: tableW - (20 + 55) * MM },
  ];
  const GRIS_ENTETE = [0.4, 0.4, 0.42];
  const headerH = 11 * MM,
    rowH = 10 * MM;
  const drawTableHeader = () => {
    doc.rect(tableX, y - headerH, tableW, headerH, GRIS_ENTETE);
    let cx = tableX;
    cols.forEach((c) => {
      doc.textCentered(cx + c.w / 2, y - headerH / 2 - 2.5, c.label, { size: 11, bold: true, color: [1, 1, 1] });
      cx += c.w;
    });
    y -= headerH;
  };
  let tableTopY = y;
  let rowBottoms = [];
  const dessinerGrille = (hautPage) => {
    doc.line(tableX, hautPage, tableX, y, 0.9, [0, 0, 0]);
    doc.line(tableX + tableW, hautPage, tableX + tableW, y, 0.9, [0, 0, 0]);
    doc.line(tableX, hautPage, tableX + tableW, hautPage, 0.9, [0, 0, 0]);
    rowBottoms.forEach((ry) => doc.line(tableX, ry, tableX + tableW, ry, 0.4, [0.55, 0.55, 0.55]));
    let cx = tableX;
    doc.line(cx, hautPage, cx, y, 0.4, [0.55, 0.55, 0.55]);
    cols.forEach((c) => {
      cx += c.w;
      doc.line(cx, hautPage, cx, y, 0.4, [0.55, 0.55, 0.55]);
    });
    doc.line(tableX, y, tableX + tableW, y, 0.9, [0, 0, 0]);
  };
  drawTableHeader();
  if (liste.length === 0) {
    doc.text(tableX + 2 * MM, y - 5 * MM, "Aucun étudiant dans ce groupe.", { size: 10, color: [0.5, 0.5, 0.5] });
    y -= 7 * MM;
  }
  liste.forEach((et, i) => {
    if (y < 35 * MM) {
      dessinerGrille(tableTopY);
      doc.newPage();
      doc.rect(0, H - 1.2 * MM, W, 1.2 * MM, ORANGE_RGB);
      y = H - 30 * MM;
      tableTopY = y;
      drawTableHeader();
      rowBottoms = [];
    }
    if (i % 2 === 1) doc.rect(tableX, y - rowH, tableW, rowH, [0.96, 0.96, 0.97]);
    const vals = [String(i + 1), et.numInscription || "—", `${et.nom || ""} ${et.prenom || ""}`.trim()];
    let cx = tableX;
    vals.forEach((v, ci) => {
      const maxW = cols[ci].w - 4 * MM;
      const texte = truncateFit(v, maxW, 11);
      if (ci === 2) doc.text(cx + 2 * MM, y - rowH / 2 - 1.8, texte, { size: 11 });
      else doc.textCentered(cx + cols[ci].w / 2, y - rowH / 2 - 1.8, texte, { size: 11 });
      cx += cols[ci].w;
    });
    y -= rowH;
    rowBottoms.push(y);
  });
  dessinerGrille(tableTopY);
  doc.textCentered(W / 2, 8 * MM, "ESRMI — École Supérieure de Rabat en Management et Ingénierie", {
    size: 8,
    color: [0.5, 0.5, 0.5],
  });
  downloadPDF(
    doc.build(),
    `Liste_Etudiants_${(filiereObj?.code || filiereObj?.intitule || "groupe").replace(/\s+/g, "_")}.pdf`
  );
}
function FichePaiementEtudiant({ ep, onClose, onSave, user, role, anneeActive, inline }) {
  const canEditPaiement = ["administrateur", "directrice"].includes(role);
  const [tab, setTab] = useState("solde");
  const [showVersement, setShowVersement] = useState(false);
  const [showRemise, setShowRemise] = useState(false);
  const [versForm, setVersForm] = useState({
    date: today(),
    montant: "",
    mode: "Virement",
    reference: "",
    banque: "",
    numCheque: "",
    motif: "Frais de formation",
    motifAutre: "",
    anneeUniversitaire: anneeActive,
  });
  const [remiseForm, setRemiseForm] = useState({
    type: "Bourse",
    pourcentage: "",
    montant: "",
    motif: "",
    valide: false,
  });
  const restant = ep.totalAPayer - ep.totalPaye;
  const taux = ep.totalAPayer > 0 ? Math.round((ep.totalPaye / ep.totalAPayer) * 100) : 0;
  const totalRemises = ep.remises.filter((r) => r.valide).reduce((s, r) => s + r.montant, 0);
  const getPayeAnnee = (anneeStr) =>
    (ep.versements || []).filter((v) => v.anneeUniversitaire === anneeStr).reduce((s, v) => s + (v.montant || 0), 0);
  const anneesAvecVersements = ANNEES_UNIVERSITAIRES.filter((a) => getPayeAnnee(a) > 0);
  const selAnneeInfo = versForm.anneeUniversitaire ? { paye: getPayeAnnee(versForm.anneeUniversitaire) } : null;
  const handleVersement = () => {
    const montant = parseFloat(versForm.montant);
    if (!montant || montant <= 0) return;
    const recu = genRecu();
    const motifFinal =
      versForm.motif === "Autre" && versForm.motifAutre ? `Autre — ${versForm.motifAutre}` : versForm.motif;
    const newVers = { ...versForm, motif: motifFinal, id: Date.now(), montant, recu, agent: user?.email || "admin" };
    delete newVers.motifAutre;
    const newEp = { ...ep, totalPaye: ep.totalPaye + montant, versements: [...ep.versements, newVers] };
    onSave(newEp, newVers);
    setShowVersement(false);
    setVersForm({
      date: today(),
      montant: "",
      mode: "Virement",
      reference: "",
      banque: "",
      numCheque: "",
      motif: "Frais de formation",
      motifAutre: "",
      anneeUniversitaire: anneeActive,
    });
  };
  const handleDeleteVersement = (v) => {
    if (
      !window.confirm(
        `Supprimer ce versement de ${formatMAD(v.montant)} (reçu ${v.recu}) ? Il sera conservé définitivement dans les archives (Paiements > Paiements archivés).`
      )
    )
      return;
    const newEp = {
      ...ep,
      totalPaye: Math.max(0, ep.totalPaye - v.montant),
      versements: ep.versements.filter((x) => x.id !== v.id),
    };
    onSave(newEp);
    if (user?.token) {
      const archiver = () =>
        saveToSP(
          "ECOGEST_VersementsArchives",
          {
            Title: v.recu || ep.numInscription || "",
            EtudiantId: String(ep.etudiantId || ""),
            NumInscription: ep.numInscription || "",
            NomEtudiant: `${ep.prenom || ""} ${ep.nom || ""}`.trim(),
            Filiere: ep.filiere || "",
            AnneeUniversitaire: v.anneeUniversitaire || "",
            DateVersement: v.date || "",
            Montant: String(v.montant || 0),
            Mode: v.mode || "",
            Reference: v.reference || "",
            Banque: v.banque || "",
            NumCheque: v.numCheque || "",
            Motif: v.motif || "",
            Recu: v.recu || "",
            Agent: v.agent || "",
            DateSuppression: today(),
            SupprimePar: user?.nom || user?.email || "Administrateur",
          },
          user.token
        );
      if (v._spId) {
        archiver().then(() => deleteFromSP("ECOGEST_Versements", v._spId, user.token));
      } else {
        archiver();
      }
    }
  };
  const handleRemise = () => {
    const montant = remiseForm.montant
      ? parseFloat(remiseForm.montant)
      : Math.round((ep.totalAPayer * parseFloat(remiseForm.pourcentage)) / 100);
    const newRemise = { ...remiseForm, id: Date.now(), montant, pourcentage: parseFloat(remiseForm.pourcentage || 0) };
    const newEp = { ...ep, remises: [...ep.remises, newRemise], totalAPayer: ep.totalAPayer - montant };
    onSave(newEp);
    setShowRemise(false);
  };
  const historiqueDerive = [
    ...ep.versements.map((v) => ({
      date: v.date,
      action: "Versement",
      montant: v.montant,
      agent: v.agent,
      detail: `${v.motif || ""} - ${v.mode} - ${v.recu} - ${v.anneeUniversitaire || ""}`,
    })),
    ...ep.remises.map((r) => ({ date: "", action: "Remise", montant: r.montant, agent: "", detail: r.motif || "" })),
    ...ep.penalites
      .filter((p) => p.applique)
      .map((p) => ({ date: p.date || "", action: "Pénalité", montant: p.montant, agent: "", detail: p.motif || "" })),
  ];
  const TABS = [
    ["solde", "💰 Solde"],
    ["versements", "💳 Versements"],
    ["remises", "🎁 Remises"],
    ["historique", "📋 Historique"],
  ];
  const body = (
    /* @__PURE__ */ <React.Fragment>
      <div
        style={{
          display: "flex",
          gap: 4,
          marginBottom: 16,
          flexWrap: "wrap",
          borderBottom: `1px solid ${C.border}`,
          paddingBottom: 8,
        }}
      >
        {TABS.map(([k, l]) => (
          /* @__PURE__ */ <button
            key={k}
            onClick={() => setTab(k)}
            style={{
              padding: "6px 12px",
              borderRadius: 8,
              border: `1px solid ${tab === k ? C.navy : C.border}`,
              background: tab === k ? C.navy : C.white,
              color: tab === k ? C.white : "#555",
              fontSize: 12,
              cursor: "pointer",
              fontFamily: "inherit",
              fontWeight: tab === k ? 600 : 400,
            }}
          >
            {l}
          </button>
        ))}
      </div>
      {tab === "solde" && (
        /* @__PURE__ */ <div>
          <div
            className="eco-grid3"
            style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 10, marginBottom: 8 }}
          >
            {[
              ["Total à payer", restant + ep.totalPaye, C.navy],
              ["Total payé", ep.totalPaye, C.green],
              ["Reste à payer", restant, restant > 0 ? C.red : C.green],
              ["Remises accordées (total)", totalRemises, C.purple],
              ["Taux de paiement", null, C.navy],
            ].map(([lbl, val, col], i) => (
              /* @__PURE__ */ <div
                key={i}
                style={{
                  background: col + "10",
                  border: `1px solid ${col}30`,
                  borderRadius: 10,
                  padding: "12px 14px",
                  borderTop: `3px solid ${col}`,
                }}
              >
                <div style={{ fontSize: 10, color: C.muted, marginBottom: 4 }}>{lbl}</div>
                {i === 4 ? (
                  /* @__PURE__ */ <div>
                    <div style={{ fontSize: 20, fontWeight: 700, color: col }}>{taux}%</div>
                    <div style={{ background: "#eee", borderRadius: 4, height: 6, marginTop: 6 }}>
                      <div
                        style={{
                          width: `${Math.min(taux, 100)}%`,
                          height: "100%",
                          background: taux >= 100 ? C.green : taux >= 50 ? C.orange : C.red,
                          borderRadius: 4,
                        }}
                      />
                    </div>
                  </div>
                ) : (
                  /* @__PURE__ */ <div style={{ fontSize: 20, fontWeight: 700, color: col }}>{formatMAD(val)}</div>
                )}
              </div>
            ))}
          </div>
          <div style={{ fontSize: 11, color: C.muted, marginBottom: 16 }}>
            💡 "Total à payer" correspond au <strong>tarif global de la formation</strong> défini dans la fiche étudiant
            (toutes années confondues) — il n'y a plus de montant dû réparti par année universitaire.
          </div>
          <div style={{ background: C.light, borderRadius: 10, padding: 14, marginBottom: 14 }}>
            <div
              style={{ fontSize: 11, fontWeight: 700, color: C.orange, textTransform: "uppercase", marginBottom: 8 }}
            >
              Informations
            </div>
            <div
              className="eco-grid2"
              style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 6, fontSize: 12 }}
            >
              {[
                ["Filière", ep.filiere],
                ["Promotion", ep.promotion],
                ["Statut", ep.statut],
                ["Responsable financier", ep.responsableFinancier || "—"],
              ].map(([l, v]) => (
                /* @__PURE__ */ <div key={l}>
                  <span style={{ color: C.muted }}>{l}: </span>
                  <strong>{v}</strong>
                </div>
              ))}
            </div>
          </div>
          <div
            style={{
              display: "flex",
              gap: 10,
              flexWrap: "wrap",
              justifyContent: "space-between",
              alignItems: "center",
            }}
          >
            <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
              {canEditPaiement && (
                /* @__PURE__ */ <Btn onClick={() => setShowVersement(true)}>💳 Enregistrer un versement</Btn>
              )}
              {canEditPaiement && (
                /* @__PURE__ */ <Btn variant="light" onClick={() => setShowRemise(true)}>
                  🎁 Accorder une remise
                </Btn>
              )}
            </div>
            {!inline && (
              /* @__PURE__ */ <button
                onClick={onClose}
                style={{
                  border: "none",
                  borderRadius: 11,
                  padding: "10px 22px",
                  cursor: "pointer",
                  fontFamily: "inherit",
                  fontSize: 13.5,
                  fontWeight: 800,
                  color: "white",
                  background: `linear-gradient(135deg, ${C.purple} 0%, ${C.teal} 100%)`,
                  boxShadow: "0 4px 12px rgba(124,58,237,0.35)",
                  transition: "transform 0.15s, box-shadow 0.15s",
                }}
                onMouseEnter={(e) => {
                  e.currentTarget.style.transform = "translateY(-2px)";
                  e.currentTarget.style.boxShadow = "0 6px 16px rgba(124,58,237,0.45)";
                }}
                onMouseLeave={(e) => {
                  e.currentTarget.style.transform = "translateY(0)";
                  e.currentTarget.style.boxShadow = "0 4px 12px rgba(124,58,237,0.35)";
                }}
              >
                ✓ Fermer
              </button>
            )}
          </div>
        </div>
      )}
      {tab === "versements" && (
        /* @__PURE__ */ <div>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12 }}>
            <div style={{ fontSize: 13, color: C.muted }}>{ep.versements.length} versement(s)</div>
            {canEditPaiement && (
              /* @__PURE__ */ <Btn small={true} onClick={() => setShowVersement(true)}>
                + Versement
              </Btn>
            )}
          </div>
          {ep.versements.length === 0 ? (
            /* @__PURE__ */ <div style={{ textAlign: "center", padding: 24, color: C.muted, fontSize: 13 }}>
              Aucun versement enregistré
            </div>
          ) : (
            [...ep.versements].reverse().map((v) => (
              /* @__PURE__ */ <div
                key={v.id}
                style={{
                  background: C.light,
                  borderRadius: 8,
                  padding: "10px 14px",
                  marginBottom: 6,
                  border: `1px solid ${C.border}`,
                }}
              >
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
                  <div>
                    <div style={{ fontWeight: 700, fontSize: 14, color: C.navy }}>{formatMAD(v.montant)}</div>
                    {v.motif && (
                      /* @__PURE__ */ <div style={{ marginTop: 2 }}>
                        <Badge color={C.orange}>{v.motif}</Badge>
                      </div>
                    )}
                    <div style={{ fontSize: 12, color: C.muted, marginTop: 4 }}>
                      {formatDateFR(v.date)} · {v.mode} · Réf: {v.reference}
                    </div>
                    {v.banque && (
                      /* @__PURE__ */ <div style={{ fontSize: 12, color: C.muted }}>
                        Banque: {v.banque} {v.numCheque ? `· Chèque: ${v.numCheque}` : ""}
                      </div>
                    )}
                    {v.anneeUniversitaire && (
                      /* @__PURE__ */ <div style={{ fontSize: 11, color: C.muted }}>
                        Année universitaire: <strong>{v.anneeUniversitaire}</strong>
                      </div>
                    )}
                    <div style={{ fontSize: 11, color: C.muted }}>Agent: {v.agent}</div>
                  </div>
                  <div style={{ textAlign: "right" }}>
                    <Badge color={C.green}>{v.recu}</Badge>
                    <div style={{ marginTop: 6, display: "flex", gap: 6, justifyContent: "flex-end" }}>
                      <Btn small={true} variant="light" onClick={() => genererRecu(ep, v)}>
                        🖨️ Reçu
                      </Btn>
                      {role === "administrateur" && (
                        /* @__PURE__ */ <Btn small={true} variant="danger" onClick={() => handleDeleteVersement(v)}>
                          🗑
                        </Btn>
                      )}
                    </div>
                  </div>
                </div>
              </div>
            ))
          )}
        </div>
      )}
      {tab === "remises" && (
        /* @__PURE__ */ <div>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12 }}>
            <div style={{ fontSize: 13, color: C.muted }}>{ep.remises.length} remise(s) accordée(s)</div>
            {canEditPaiement && (
              /* @__PURE__ */ <Btn small={true} onClick={() => setShowRemise(true)}>
                + Remise
              </Btn>
            )}
          </div>
          {ep.remises.map((r) => (
            /* @__PURE__ */ <div
              key={r.id}
              style={{
                background: r.valide ? "#f0fff4" : "#fff5f5",
                borderRadius: 8,
                padding: "10px 14px",
                marginBottom: 6,
                border: `1px solid ${r.valide ? "#9ae6b4" : "#fecaca"}`,
              }}
            >
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                <div>
                  <div style={{ fontWeight: 600, fontSize: 13, color: C.navy }}>{r.type}</div>
                  <div style={{ fontSize: 12, color: C.muted }}>{r.motif}</div>
                  {r.pourcentage > 0 && (
                    /* @__PURE__ */ <div style={{ fontSize: 12, color: C.muted }}>{r.pourcentage}% de réduction</div>
                  )}
                </div>
                <div style={{ textAlign: "right" }}>
                  <div style={{ fontWeight: 700, fontSize: 15, color: r.valide ? C.green : C.orange }}>
                    -{formatMAD(r.montant)}
                  </div>
                  <Badge color={r.valide ? C.green : C.orange}>{r.valide ? "Validée" : "En attente"}</Badge>
                </div>
              </div>
            </div>
          ))}
          {ep.remises.length === 0 && (
            /* @__PURE__ */ <div style={{ textAlign: "center", padding: 24, color: C.muted, fontSize: 13 }}>
              Aucune remise
            </div>
          )}
        </div>
      )}
      {tab === "historique" && (
        /* @__PURE__ */ <div>
          {historiqueDerive.length === 0 ? (
            /* @__PURE__ */ <div style={{ textAlign: "center", padding: 24, color: C.muted, fontSize: 13 }}>
              Aucun historique
            </div>
          ) : (
            [...historiqueDerive].reverse().map((h, i) => (
              /* @__PURE__ */ <div
                key={i}
                style={{ display: "flex", gap: 12, padding: "8px 0", borderBottom: `1px solid ${C.light}` }}
              >
                <div
                  style={{
                    width: 36,
                    height: 36,
                    borderRadius: 8,
                    background: C.light,
                    display: "flex",
                    alignItems: "center",
                    justifyContent: "center",
                    fontSize: 18,
                    flexShrink: 0,
                  }}
                >
                  {h.action === "Versement" ? "💳" : h.action === "Remise" ? "🎁" : "📝"}
                </div>
                <div>
                  <div style={{ fontWeight: 600, fontSize: 13 }}>{h.action}</div>
                  <div style={{ fontSize: 12, color: C.muted }}>
                    {formatDateFR(h.date)} · {h.agent}
                  </div>
                  <div style={{ fontSize: 12, color: "#555" }}>{h.detail}</div>
                </div>
                {h.montant && (
                  /* @__PURE__ */ <div style={{ marginLeft: "auto", fontWeight: 700, color: C.navy, fontSize: 14 }}>
                    {formatMAD(h.montant)}
                  </div>
                )}
              </div>
            ))
          )}
        </div>
      )}
      {showVersement && (
        /* @__PURE__ */ <Modal title="Enregistrer un versement" onClose={() => setShowVersement(false)} width={480}>
          <div
            style={{
              background: C.navyLight,
              borderRadius: 8,
              padding: "8px 14px",
              fontSize: 12,
              color: C.navy,
              marginBottom: 14,
            }}
          >
            📅 Année universitaire active : <strong>{anneeActive}</strong> — appliquée automatiquement (réglable dans
            Paramétrage)
          </div>
          <div
            style={{
              background: "#f0f7ff",
              border: "1px solid #bee3f8",
              borderRadius: 8,
              padding: "10px 14px",
              fontSize: 12,
              color: "#2b6cb0",
              marginBottom: 14,
            }}
          >
            💰 Tarif global de la formation : <strong>{formatMAD(ep.totalAPayer)}</strong> · Déjà payé (tout confondu) : <strong>{formatMAD(ep.totalPaye)}</strong> · Reste : <strong style={{ color: restant > 0 ? C.red : C.green }}>{formatMAD(restant)}</strong>
            {selAnneeInfo && (
              /* @__PURE__ */ <div style={{ marginTop: 4 }}>
                Déjà payé pour <strong>{anneeActive}</strong> : <strong>{formatMAD(selAnneeInfo.paye)}</strong>
              </div>
            )}
          </div>
          <div>
            <div className="eco-grid2" style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "0 14px" }}>
              <Input
                label="Date *"
                type="date"
                value={versForm.date}
                onChange={(e) => setVersForm((f) => ({ ...f, date: e.target.value }))}
              />
              <Input
                label="Montant (MAD) *"
                type="number"
                value={versForm.montant}
                onChange={(e) => setVersForm((f) => ({ ...f, montant: e.target.value }))}
                placeholder="0.00"
              />
            </div>
            <Sel
              label="Mode de paiement *"
              value={versForm.mode}
              onChange={(e) => setVersForm((f) => ({ ...f, mode: e.target.value }))}
            >
              {["Espèces", "Chèque", "Virement", "Carte bancaire", "Paiement en ligne"].map((m) => (
                /* @__PURE__ */ <option key={m}>{m}</option>
              ))}
            </Sel>
            <Sel
              label="Motif du paiement *"
              value={versForm.motif}
              onChange={(e) => setVersForm((f) => ({ ...f, motif: e.target.value }))}
            >
              {MOTIFS_PAIEMENT.map((m) => (
                /* @__PURE__ */ <option key={m} value={m}>
                  {m}
                </option>
              ))}
            </Sel>
            {versForm.motif === "Autre" && (
              /* @__PURE__ */ <Input
                label="Précision du motif"
                value={versForm.motifAutre}
                onChange={(e) => setVersForm((f) => ({ ...f, motifAutre: e.target.value }))}
                placeholder="Préciser le motif..."
              />
            )}
            <div className="eco-grid2" style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "0 14px" }}>
              <Input
                label="Référence"
                value={versForm.reference}
                onChange={(e) => setVersForm((f) => ({ ...f, reference: e.target.value }))}
                placeholder="VIR001..."
              />
              <Input
                label="Banque"
                value={versForm.banque}
                onChange={(e) => setVersForm((f) => ({ ...f, banque: e.target.value }))}
                placeholder="Attijariwafa..."
              />
            </div>
            {versForm.mode === "Chèque" && (
              /* @__PURE__ */ <Input
                label="N° Chèque"
                value={versForm.numCheque}
                onChange={(e) => setVersForm((f) => ({ ...f, numCheque: e.target.value }))}
              />
            )}
          </div>
          <div
            style={{
              background: "#f0fff4",
              border: "1px solid #9ae6b4",
              borderRadius: 8,
              padding: "8px 14px",
              fontSize: 12,
              color: C.green,
              marginBottom: 12,
            }}
          >
            ✅ Un reçu sera généré automatiquement après enregistrement
          </div>
          <div style={{ display: "flex", gap: 10, justifyContent: "flex-end" }}>
            <Btn variant="light" onClick={() => setShowVersement(false)}>
              Annuler
            </Btn>
            <Btn onClick={handleVersement} disabled={!versForm.montant || !versForm.motif}>
              💳 Enregistrer & Générer reçu
            </Btn>
          </div>
        </Modal>
      )}
      {showRemise && (
        /* @__PURE__ */ <Modal title="Accorder une remise" onClose={() => setShowRemise(false)} width={440}>
          <Sel
            label="Type de remise *"
            value={remiseForm.type}
            onChange={(e) => setRemiseForm((f) => ({ ...f, type: e.target.value }))}
          >
            {["Bourse", "Réduction familiale", "Réduction exceptionnelle", "Remise partenaire", "Remise personnel"].map(
              (t) => (
                /* @__PURE__ */ <option key={t}>{t}</option>
              )
            )}
          </Sel>
          <div className="eco-grid2" style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "0 14px" }}>
            <Input
              label="Pourcentage (%)"
              type="number"
              min="0"
              max="100"
              value={remiseForm.pourcentage}
              onChange={(e) => setRemiseForm((f) => ({ ...f, pourcentage: e.target.value }))}
            />
            <Input
              label="Montant fixe (MAD)"
              type="number"
              value={remiseForm.montant}
              onChange={(e) => setRemiseForm((f) => ({ ...f, montant: e.target.value }))}
            />
          </div>
          <Input
            label="Motif *"
            value={remiseForm.motif}
            onChange={(e) => setRemiseForm((f) => ({ ...f, motif: e.target.value }))}
          />
          <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 14 }}>
            <input
              type="checkbox"
              checked={remiseForm.valide}
              onChange={(e) => setRemiseForm((f) => ({ ...f, valide: e.target.checked }))}
              id="valide"
            />
            <label htmlFor="valide" style={{ fontSize: 13, color: "#555" }}>
              Validé par la direction
            </label>
          </div>
          <div style={{ display: "flex", gap: 10, justifyContent: "flex-end" }}>
            <Btn variant="light" onClick={() => setShowRemise(false)}>
              Annuler
            </Btn>
            <Btn onClick={handleRemise} disabled={!remiseForm.motif}>
              🎁 Appliquer la remise
            </Btn>
          </div>
        </Modal>
      )}
    </React.Fragment>
  );
  if (inline) {
    return (
      /* @__PURE__ */ <Card style={{ padding: 28 }}>
        <h3 style={{ margin: "0 0 20px", color: C.navy, fontSize: 17, fontWeight: 700, letterSpacing: "-0.01em" }}>
          Paiement — {ep.prenom} {ep.nom} · {ep.numInscription}
        </h3>
        {body}
      </Card>
    );
  }
  return (
    /* @__PURE__ */ <Modal
      title={`Paiement — ${ep.prenom} ${ep.nom} · ${ep.numInscription}`}
      onClose={onClose}
      width={720}
      hideFooterClose={true}
    >
      {body}
    </Modal>
  );
}
function Paiements({ data, setData, role, user, toast }) {
  const [tab, setTab] = useState("dashboard");
  const [search, setSearch] = useState("");
  const [showFiche, setShowFiche] = useState(null);
  const [editVisiteur, setEditVisiteur] = useState(null);
  const [filterStatut, setFilterStatut] = useState("");
  const [filterFiliere, setFilterFiliere] = useState("");
  const pData = data.paiements || {
    etudiants_paiement: [],
    parametres: DEMO_PAIEMENT.parametres,
    parametres_penalites: DEMO_PAIEMENT.parametres_penalites,
    parametres_alertes: DEMO_PAIEMENT.parametres_alertes,
  };
  const paiementsExistants = pData.etudiants_paiement || [];
  const eps = (data.etudiants || []).map((et) => {
    const conventionFields = {
      tarifGlobalFormation: et.tarifGlobalFormation || 0,
      responsableFinancier: et.responsableFinancier || "",
    };
    const totalConvention = conventionFields.tarifGlobalFormation;
    const doublons = paiementsExistants.filter(
      (e) => (et.numInscription && e.numInscription === et.numInscription) || e.etudiantId === et.id
    );
    const existing =
      doublons.length > 1
        ? doublons.sort(
            (a, b) =>
              (b.versements?.length || 0) - (a.versements?.length || 0) || (b.totalPaye || 0) - (a.totalPaye || 0)
          )[0]
        : doublons[0];
    if (existing) {
      const remisesTotal = (existing.remises || []).reduce((s, r) => s + (r.montant || 0), 0);
      return { ...existing, ...conventionFields, totalAPayer: totalConvention - remisesTotal };
    }
    return {
      id: et.id,
      _spId: null,
      etudiantId: et.id,
      nom: et.nom || "",
      prenom: et.prenom || "",
      numInscription: et.numInscription || "",
      filiere: et.filiere || "",
      niveau: et.niveau || "",
      promotion: et.annee || et.promotion || "",
      statut: et.statut || "Actif",
      cin: et.cin || "",
      telephone: et.telephone || "",
      email: et.email || "",
      responsableFinancier: et.responsableFinancier || "",
      ...conventionFields,
      totalAPayer: totalConvention,
      totalPaye: 0,
      remises: [],
      versements: [],
      penalites: [],
      historique: [],
    };
  });
  const updateEtudiantPaiement = (newEp, newVersement) => {
    if (!user?.token) {
      toast("⚠️ Non connecté à SharePoint — le paiement n'a pas été enregistré");
      return;
    }
    const statutPaiement =
      (newEp.totalAPayer || 0) > 0 && (newEp.totalPaye || 0) >= (newEp.totalAPayer || 0) ? "Soldé" : "Régulier";
    const applyLocalUpdate = (finalEp) => {
      const exists = paiementsExistants.some((e) => e.id === finalEp.id);
      const newEps = exists
        ? paiementsExistants.map((e) => (e.id === finalEp.id ? finalEp : e))
        : [...paiementsExistants, finalEp];
      setData((d) => ({ ...d, paiements: { ...pData, etudiants_paiement: newEps } }));
      if (showFiche?.id === finalEp.id) setShowFiche(finalEp);
    };
    const saveSummary = (ep) =>
      saveToSP(
        "ECOGEST_Paiements",
        {
          Title: ep.numInscription || "",
          EtudiantId: String(ep.etudiantId || ""),
          NumInscription: ep.numInscription || "",
          NomEtudiant: `${ep.prenom || ""} ${ep.nom || ""}`.trim(),
          Filiere: ep.filiere || "",
          AnneeScolaire: ep.promotion || "",
          TotalAPayer: String(ep.totalAPayer || 0),
          TotalPaye: String(ep.totalPaye || 0),
          Solde: String((ep.totalAPayer || 0) - (ep.totalPaye || 0)),
          Statut: statutPaiement,
          RemisesJSON: JSON.stringify(ep.remises || []),
          PenalitesJSON: JSON.stringify(ep.penalites || []),
        },
        user.token,
        ep._spId
      ).then((result) => {
        if (!result) {
          toast("⚠️ Échec de la mise à jour du résumé dans SharePoint — vérifiez la console (F12)");
          return null;
        }
        if (result?.id && !ep._spId) {
          setData((d) => ({
            ...d,
            paiements: {
              ...d.paiements,
              etudiants_paiement: (d.paiements?.etudiants_paiement || []).map((e) =>
                e.id === ep.id ? { ...e, _spId: result.id } : e
              ),
            },
          }));
          setShowFiche((sf) => (sf && sf.id === ep.id ? { ...sf, _spId: result.id } : sf));
        }
        return result;
      });
    if (newVersement) {
      saveToSP(
        "ECOGEST_Versements",
        {
          Title: newVersement.recu || newEp.numInscription || "",
          EtudiantId: String(newEp.etudiantId || ""),
          NumInscription: newEp.numInscription || "",
          NomEtudiant: `${newEp.prenom || ""} ${newEp.nom || ""}`.trim(),
          Filiere: newEp.filiere || "",
          AnneeUniversitaire: newVersement.anneeUniversitaire || "",
          DateVersement: newVersement.date || "",
          Montant: String(newVersement.montant || 0),
          Mode: newVersement.mode || "",
          Reference: newVersement.reference || "",
          Banque: newVersement.banque || "",
          NumCheque: newVersement.numCheque || "",
          Motif: newVersement.motif || "",
          Recu: newVersement.recu || "",
          Agent: newVersement.agent || "",
        },
        user.token,
        null
      ).then((result) => {
        if (!result) {
          toast("⚠️ Le versement n'a pas pu être enregistré dans SharePoint — il ne s'affiche pas");
          return;
        }
        const confirmedVersement = { ...newVersement, _spId: result.id };
        const finalEp = {
          ...newEp,
          versements: newEp.versements.map((v) => (v.id === newVersement.id ? confirmedVersement : v)),
        };
        applyLocalUpdate(finalEp);
        saveSummary(finalEp);
        toast("Paiement enregistré dans SharePoint ✓");
        setTimeout(() => genererRecu(finalEp, confirmedVersement), 200);
      });
    } else {
      saveSummary(newEp).then((result) => {
        if (!result) return;
        applyLocalUpdate(newEp);
        toast("Paiement mis à jour dans SharePoint ✓");
      });
    }
  };
  const handleRestaurerVersement = (v) => {
    if (!user?.token) {
      toast("⚠️ Non connecté à SharePoint");
      return;
    }
    const ep = eps.find((e) => e.numInscription === v.numInscription);
    if (!ep) {
      toast(
        `⚠️ Aucun compte de paiement actif trouvé pour ${v.nomEtudiant} (${v.numInscription}) — restauration impossible`
      );
      return;
    }
    if (
      !window.confirm(
        `Restaurer le versement de ${formatMAD(v.montant)} (reçu ${v.recu}) sur le compte de ${ep.prenom} ${ep.nom} ?`
      )
    )
      return;
    const versementRestaure = {
      id: Date.now(),
      date: v.date,
      montant: v.montant,
      mode: v.mode,
      reference: v.reference,
      banque: v.banque,
      numCheque: v.numCheque,
      motif: v.motif,
      recu: v.recu,
      agent: v.agent,
      anneeUniversitaire: v.anneeUniversitaire,
    };
    saveToSP(
      "ECOGEST_Versements",
      {
        Title: v.recu || ep.numInscription || "",
        EtudiantId: String(ep.etudiantId || ""),
        NumInscription: ep.numInscription || "",
        NomEtudiant: `${ep.prenom || ""} ${ep.nom || ""}`.trim(),
        Filiere: ep.filiere || "",
        AnneeUniversitaire: v.anneeUniversitaire || "",
        DateVersement: v.date || "",
        Montant: String(v.montant || 0),
        Mode: v.mode || "",
        Reference: v.reference || "",
        Banque: v.banque || "",
        NumCheque: v.numCheque || "",
        Motif: v.motif || "",
        Recu: v.recu || "",
        Agent: v.agent || "",
      },
      user.token,
      null
    ).then((result) => {
      if (!result) {
        toast("⚠️ Échec de la restauration du versement dans SharePoint");
        return;
      }
      const confirmedVersement = { ...versementRestaure, _spId: result.id };
      const finalEp = {
        ...ep,
        totalPaye: (ep.totalPaye || 0) + v.montant,
        versements: [...ep.versements, confirmedVersement],
      };
      const exists = paiementsExistants.some((e) => e.id === finalEp.id);
      const newEps = exists
        ? paiementsExistants.map((e) => (e.id === finalEp.id ? finalEp : e))
        : [...paiementsExistants, finalEp];
      setData((d) => ({
        ...d,
        paiements: { ...pData, etudiants_paiement: newEps },
        versementsArchives: (d.versementsArchives || []).filter((a) => a.id !== v.id),
      }));
      saveToSP(
        "ECOGEST_Paiements",
        {
          Title: finalEp.numInscription || "",
          EtudiantId: String(finalEp.etudiantId || ""),
          NumInscription: finalEp.numInscription || "",
          NomEtudiant: `${finalEp.prenom || ""} ${finalEp.nom || ""}`.trim(),
          Filiere: finalEp.filiere || "",
          AnneeScolaire: finalEp.promotion || "",
          TotalAPayer: String(finalEp.totalAPayer || 0),
          TotalPaye: String(finalEp.totalPaye || 0),
          Solde: String((finalEp.totalAPayer || 0) - (finalEp.totalPaye || 0)),
          Statut: (finalEp.totalAPayer || 0) > 0 && finalEp.totalPaye >= finalEp.totalAPayer ? "Soldé" : "Régulier",
          RemisesJSON: JSON.stringify(finalEp.remises || []),
          PenalitesJSON: JSON.stringify(finalEp.penalites || []),
        },
        user.token,
        finalEp._spId
      );
      if (v._spId) deleteFromSP("ECOGEST_VersementsArchives", v._spId, user.token);
      toast(`✅ Versement de ${formatMAD(v.montant)} restauré sur le compte de ${ep.prenom} ${ep.nom}`);
    });
  };
  const handleRestaurerPaiement = (p) => {
    if (!user?.token) {
      toast("⚠️ Non connecté à SharePoint");
      return;
    }
    const dejaExistant = eps.find((e) => e.numInscription === p.numInscription && e._spId);
    if (dejaExistant) {
      toast(
        `⚠️ Un compte de paiement actif existe déjà pour ${p.nomEtudiant} — restauration annulée pour éviter un doublon`
      );
      return;
    }
    if (!window.confirm(`Restaurer le compte de paiement de ${p.nomEtudiant} (${p.numInscription}) ?`)) return;
    const et = (data.etudiants || []).find((e) => e.numInscription === p.numInscription);
    const nomComplet = (p.nomEtudiant || "").split(" ");
    const finalEp = {
      id: et?.id || Date.now(),
      _spId: null,
      etudiantId: et?.id || "",
      nom: et?.nom || nomComplet.slice(1).join(" "),
      prenom: et?.prenom || nomComplet[0] || "",
      numInscription: p.numInscription,
      filiere: p.filiere,
      niveau: et?.niveau || "",
      promotion: p.promotion,
      statut: "Actif",
      cin: et?.cin || "",
      telephone: et?.telephone || "",
      email: et?.email || "",
      responsableFinancier: et?.responsableFinancier || "",
      tarifGlobalFormation: et?.tarifGlobalFormation || p.totalAPayer || 0,
      totalAPayer: p.totalAPayer || 0,
      totalPaye: p.totalPaye || 0,
      remises: [],
      versements: [],
      penalites: [],
      historique: [],
    };
    saveToSP(
      "ECOGEST_Paiements",
      {
        Title: finalEp.numInscription || "",
        EtudiantId: String(finalEp.etudiantId || ""),
        NumInscription: finalEp.numInscription || "",
        NomEtudiant: p.nomEtudiant || "",
        Filiere: finalEp.filiere || "",
        AnneeScolaire: finalEp.promotion || "",
        TotalAPayer: String(finalEp.totalAPayer || 0),
        TotalPaye: String(finalEp.totalPaye || 0),
        Solde: String((finalEp.totalAPayer || 0) - (finalEp.totalPaye || 0)),
        Statut: p.statut || "Régulier",
        RemisesJSON: "[]",
        PenalitesJSON: "[]",
      },
      user.token,
      null
    ).then((result) => {
      if (!result) {
        toast("⚠️ Échec de la restauration du compte de paiement dans SharePoint");
        return;
      }
      const confirmedEp = { ...finalEp, _spId: result.id };
      setData((d) => ({
        ...d,
        paiements: { ...pData, etudiants_paiement: [...paiementsExistants, confirmedEp] },
        paiementsArchives: (d.paiementsArchives || []).filter((a) => a.id !== p.id),
      }));
      if (p._spId) deleteFromSP("ECOGEST_PaiementsArchives", p._spId, user.token);
      toast(`✅ Compte de paiement de ${p.nomEtudiant} restauré`);
    });
  };
  const totalEncaisseJour = eps.reduce(
    (s, e) => s + e.versements.filter((v) => v.date === today()).reduce((a, v) => a + v.montant, 0),
    0
  );
  const totalEncaisseMois = eps.reduce(
    (s, e) =>
      s +
      e.versements
        .filter((v) => v.date?.startsWith(/* @__PURE__ */ new Date().toISOString().slice(0, 7)))
        .reduce((a, v) => a + v.montant, 0),
    0
  );
  const totalRestant = eps.reduce((s, e) => s + (e.totalAPayer - e.totalPaye), 0);
  const nbSoldeRestant = eps.filter((e) => e.totalAPayer - e.totalPaye > 0 || e.totalAPayer === 0).length;
  const totalAPayer = eps.reduce((s, e) => s + e.totalAPayer, 0);
  const totalPaye = eps.reduce((s, e) => s + e.totalPaye, 0);
  const tauxRecouvrement = totalAPayer > 0 ? Math.round((totalPaye / totalAPayer) * 100) : 0;
  const [dateDebut, setDateDebut] = useState("");
  const [dateFin, setDateFin] = useState("");
  const [filterMotif, setFilterMotif] = useState("");
  const [searchOps, setSearchOps] = useState("");
  const allVersements = eps.flatMap((e) =>
    (e.versements || []).map((v) => ({
      ...v,
      etudiantIdRef: e.id,
      etudiantNom: e.nom,
      etudiantPrenom: e.prenom,
      numInscription: e.numInscription,
      filiere: e.filiere,
    }))
  );
  const opsFiltered = allVersements
    .filter((v) => {
      const matchDebut = !dateDebut || (v.date || "") >= dateDebut;
      const matchFin = !dateFin || (v.date || "") <= dateFin;
      const matchMotif = !filterMotif || v.motif === filterMotif;
      const matchFiliere = !filterFiliere || v.filiere === filterFiliere;
      const q = searchOps.toLowerCase();
      const matchSearch =
        !q || `${v.etudiantPrenom} ${v.etudiantNom} ${v.numInscription} ${v.reference || ""}`.toLowerCase().includes(q);
      return matchDebut && matchFin && matchMotif && matchFiliere && matchSearch;
    })
    .sort((a, b) => (b.id || 0) - (a.id || 0));
  const totalOpsFiltered = opsFiltered.reduce((s, v) => s + (v.montant || 0), 0);
  const filtered = eps.filter((e) => {
    const q = search.toLowerCase();
    const match =
      !q || `${e.nom} ${e.prenom} ${e.cin || ""} ${e.numInscription} ${e.telephone || ""}`.toLowerCase().includes(q);
    const statMatch =
      !filterStatut ||
      (filterStatut === "solde"
        ? e.totalAPayer - e.totalPaye > 0
        : filterStatut === "jourpaye"
          ? e.versements.some((v) => v.date === today())
          : true);
    const filiereMatch = !filterFiliere || e.filiere === filterFiliere;
    return match && statMatch && filiereMatch;
  });
  const TABS = [
    ["dashboard", "📊 Tableau de bord"],
    ["liste", "👥 Étudiants"],
    ["operations", "🧾 Toutes les opérations"],
    ["etats", "📈 États financiers"],
    ...(role === "administrateur" ? [["archives", "🗄️ Paiements archivés"]] : []),
  ];
  if (role === "etudiant") {
    const monEp = eps.find(
      (e) =>
        (user?.numInscription && e.numInscription === user.numInscription) ||
        (user?.etudiantId && e.etudiantId === user.etudiantId)
    );
    return (
      /* @__PURE__ */ <div>
        <PageHeader title="💰 Mes paiements" sub="Consultez ici le détail de votre situation financière" />
        {monEp ? (
          /* @__PURE__ */ <FichePaiementEtudiant
            ep={monEp}
            onClose={() => {}}
            inline={true}
            onSave={updateEtudiantPaiement}
            user={user}
            role={role}
            anneeActive={data.configuration?.anneeUniversitaireActive || ANNEES_UNIVERSITAIRES[0]}
          />
        ) : (
          /* @__PURE__ */ <Card style={{ padding: 32, textAlign: "center", color: C.muted, fontSize: 13 }}>
            Aucune information de paiement n'a encore été trouvée pour votre compte.
          </Card>
        )}
      </div>
    );
  }
  return (
    /* @__PURE__ */ <div>
      <PageHeader
        title="💰 Gestion des Paiements"
        sub="Module financier — Accès restreint · données SharePoint ECOGEST"
      />
      <div
        style={{ display: "flex", gap: 6, marginBottom: 20, borderBottom: `1px solid ${C.border}`, paddingBottom: 10 }}
      >
        {TABS.map(([k, l]) => (
          /* @__PURE__ */ <button
            key={k}
            onClick={() => setTab(k)}
            style={{
              padding: "8px 16px",
              borderRadius: 8,
              border: `1px solid ${tab === k ? C.navy : C.border}`,
              background: tab === k ? C.navy : C.white,
              color: tab === k ? C.white : "#555",
              fontSize: 13,
              cursor: "pointer",
              fontFamily: "inherit",
              fontWeight: tab === k ? 600 : 400,
            }}
          >
            {l}
          </button>
        ))}
      </div>
      {tab === "dashboard" && (
        /* @__PURE__ */ <div>
          <div style={{ display: "grid", gridTemplateColumns: "repeat(3, 1fr)", gap: 12, marginBottom: 20 }}>
            <StatCard label="Montant global" value={formatMAD(totalAPayer)} color={C.navy} icon="🏦" />
            <StatCard label="Total encaissé" value={formatMAD(totalPaye)} color={C.green} icon="💰" />
            <StatCard label="Total restant" value={formatMAD(totalRestant)} color={C.orange} icon="⏳" />
            <StatCard label="Recette du jour" value={formatMAD(totalEncaisseJour)} color={C.green} icon="💵" />
            <StatCard label="Encaissement du mois" value={formatMAD(totalEncaisseMois)} color={C.navy} icon="📅" />
            <StatCard label="Taux de recouvrement" value={`${tauxRecouvrement}%`} color={C.purple} icon="📈" />
          </div>
          <Card style={{ padding: 20, marginBottom: 16 }}>
            <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 8 }}>
              <span style={{ fontWeight: 600, color: C.navy }}>Recouvrement global</span>
              <span style={{ fontWeight: 700, color: C.navy }}>
                {formatMAD(totalPaye)} / {formatMAD(totalAPayer)}
              </span>
            </div>
            <div style={{ background: "#eee", borderRadius: 8, height: 12 }}>
              <div
                style={{
                  width: `${Math.min(tauxRecouvrement, 100)}%`,
                  height: "100%",
                  background: `linear-gradient(90deg, ${C.green}, ${C.navy})`,
                  borderRadius: 8,
                  transition: "width 0.5s",
                }}
              />
            </div>
            <div
              style={{ display: "flex", justifyContent: "space-between", marginTop: 6, fontSize: 12, color: C.muted }}
            >
              <span>{tauxRecouvrement}% encaissé</span>
              <span>Reste: {formatMAD(totalRestant)}</span>
            </div>
          </Card>
        </div>
      )}
      {tab === "liste" && (
        /* @__PURE__ */ <div>
          <div style={{ display: "flex", gap: 10, marginBottom: 14, flexWrap: "wrap" }}>
            <input
              placeholder="🔍 Nom, CIN, N° inscription, téléphone..."
              value={search}
              onChange={(e) => setSearch(e.target.value)}
              style={{
                flex: 1,
                minWidth: 200,
                padding: "9px 14px",
                border: `1px solid ${C.border}`,
                borderRadius: 8,
                fontSize: 13,
                fontFamily: "inherit",
              }}
            />
            <select
              value={filterFiliere}
              onChange={(e) => setFilterFiliere(e.target.value)}
              style={{
                padding: "9px 14px",
                border: `1px solid ${C.border}`,
                borderRadius: 8,
                fontSize: 13,
                fontFamily: "inherit",
                background: C.white,
              }}
            >
              <option value="">-- Tous les groupes --</option>
              {getFilieres(data).map((f) => (
                /* @__PURE__ */ <option key={f} value={f}>
                  {f}
                </option>
              ))}
            </select>
            {[
              ["", "Tous"],
              ["solde", "Solde restant"],
              ["jourpaye", "Payé aujourd'hui"],
            ].map(([v, l]) => (
              /* @__PURE__ */ <button
                key={v}
                onClick={() => setFilterStatut(v)}
                style={{
                  padding: "8px 14px",
                  borderRadius: 20,
                  border: `1px solid ${C.border}`,
                  background: filterStatut === v ? C.navy : C.white,
                  color: filterStatut === v ? C.white : "#555",
                  fontSize: 12,
                  cursor: "pointer",
                  fontFamily: "inherit",
                }}
              >
                {l}
              </button>
            ))}
          </div>
          {(() => {
            const cols = [
              {
                key: "_filiereCode",
                label: "Filière",
                render: (_, row) => (
                  /* @__PURE__ */ <FiliereBadge code={filiereCode(row.filiere, data.filieres) || row.filiere} />
                ),
              },
              {
                key: "numInscription",
                label: "N° Inscription",
                render: (v, row) => {
                  const coulFiliere = colorForFiliereCode(filiereCode(row.filiere, data.filieres) || row.filiere);
                  return (
                    /* @__PURE__ */ <div
                      style={{ display: "flex", alignItems: "center", gap: 8, whiteSpace: "nowrap" }}
                    >
                      <span style={{ whiteSpace: "nowrap" }}>{v}</span>
                      <span
                        title={`${(row.versements || []).length} versement(s) effectué(s)`}
                        style={{
                          display: "inline-flex",
                          alignItems: "center",
                          gap: 6,
                          background: coulFiliere,
                          color: "#ffffff",
                          fontSize: 11,
                          fontWeight: 700,
                          padding: "3px 10px 3px 4px",
                          borderRadius: 20,
                          flexShrink: 0,
                          whiteSpace: "nowrap",
                          boxShadow: `0 3px 8px ${coulFiliere}80`,
                        }}
                      >
                        <span
                          style={{
                            background: "rgba(255,255,255,0.94)",
                            color: coulFiliere,
                            borderRadius: "50%",
                            width: 19,
                            height: 19,
                            display: "flex",
                            alignItems: "center",
                            justifyContent: "center",
                            fontSize: 11.5,
                            fontWeight: 900,
                            flexShrink: 0,
                          }}
                        >
                          {(row.versements || []).length}
                        </span>
                        💳 versement{(row.versements || []).length > 1 ? "s" : ""}
                      </span>
                    </div>
                  );
                },
              },
              {
                key: "nom",
                label: "Nom",
                render: (v, row) => (
                  /* @__PURE__ */ <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
                    <span>{v}</span>
                    {estEtudiantTest(row.prenom, row.nom) && (
                      /* @__PURE__ */ <span
                        title="Étudiant de test — n'existe pas réellement"
                        style={{
                          background: C.purple,
                          color: "white",
                          fontSize: 9.5,
                          fontWeight: 800,
                          borderRadius: 20,
                          padding: "1px 7px",
                          flexShrink: 0,
                          display: "inline-flex",
                          alignItems: "center",
                          gap: 3,
                        }}
                      >
                        🧪 TEST
                      </span>
                    )}
                  </div>
                ),
              },
              { key: "prenom", label: "Prénom" },
              {
                key: "id",
                label: "Solde",
                headerAlign: "center",
                render: (_, row) => /* @__PURE__ */ <SoldeBadge ep={row} />,
              },
              {
                key: "id2",
                label: "",
                render: (_, row) => (
                  /* @__PURE__ */ <Btn small={true} variant="navy" onClick={() => setShowFiche(row)}>
                    💰 Paiement
                  </Btn>
                ),
              },
            ];
            const filteredSorted = [...filtered].sort((a, b) => {
              const codeA = filiereCode(a.filiere, data.filieres) || a.filiere || "";
              const codeB = filiereCode(b.filiere, data.filieres) || b.filiere || "";
              return (
                codeA.localeCompare(codeB, "fr", { numeric: true, sensitivity: "base" }) ||
                (a.nom || "").localeCompare(b.nom || "", "fr", { sensitivity: "base" })
              );
            });
            const bandeFiliereParId = {};
            const premiereLigneFiliereId = {};
            const derniereLigneFiliereId = {};
            {
              let bande = 0,
                filierePrecedente = null;
              filteredSorted.forEach((e, idx) => {
                if (e.filiere !== filierePrecedente) {
                  bande = 1 - bande;
                  filierePrecedente = e.filiere;
                  premiereLigneFiliereId[e.id] = true;
                  if (idx > 0) derniereLigneFiliereId[filteredSorted[idx - 1].id] = true;
                }
                bandeFiliereParId[e.id] = bande;
              });
              if (filteredSorted.length > 0)
                derniereLigneFiliereId[filteredSorted[filteredSorted.length - 1].id] = true;
            }
            if (filteredSorted.length === 0)
              return (
                /* @__PURE__ */ <Card style={{ padding: 32, textAlign: "center", color: C.muted, fontSize: 13 }}>
                  Aucun étudiant trouvé
                </Card>
              );
            return (
              /* @__PURE__ */ <Card style={{ overflow: "hidden" }}>
                <DataTable
                  cols={cols}
                  rows={filteredSorted}
                  compact={true}
                  emptyMsg="Aucun étudiant trouvé"
                  rowStyle={(row) => ({
                    background: bandeFiliereParId[row.id] === 0 ? C.navyLight : "#e0f4f2",
                    borderTop: premiereLigneFiliereId[row.id] ? `2.5px solid ${C.navy}` : void 0,
                    borderBottom: derniereLigneFiliereId[row.id] ? `2.5px solid ${C.navy}` : `1px solid ${C.border}`,
                  })}
                />
              </Card>
            );
          })()}
        </div>
      )}
      {tab === "operations" && (
        /* @__PURE__ */ <div>
          <div style={{ display: "flex", gap: 10, marginBottom: 14, flexWrap: "wrap", alignItems: "flex-end" }}>
            <div>
              <label style={{ fontSize: 11, fontWeight: 600, color: "#555", display: "block", marginBottom: 4 }}>
                Du
              </label>
              <input
                type="date"
                value={dateDebut}
                onChange={(e) => setDateDebut(e.target.value)}
                style={{
                  padding: "8px 11px",
                  border: `1px solid ${C.border}`,
                  borderRadius: 8,
                  fontSize: 13,
                  fontFamily: "inherit",
                }}
              />
            </div>
            <div>
              <label style={{ fontSize: 11, fontWeight: 600, color: "#555", display: "block", marginBottom: 4 }}>
                Au
              </label>
              <input
                type="date"
                value={dateFin}
                onChange={(e) => setDateFin(e.target.value)}
                style={{
                  padding: "8px 11px",
                  border: `1px solid ${C.border}`,
                  borderRadius: 8,
                  fontSize: 13,
                  fontFamily: "inherit",
                }}
              />
            </div>
            <div>
              <label style={{ fontSize: 11, fontWeight: 600, color: "#555", display: "block", marginBottom: 4 }}>
                Motif
              </label>
              <select
                value={filterMotif}
                onChange={(e) => setFilterMotif(e.target.value)}
                style={{
                  padding: "8px 11px",
                  border: `1px solid ${C.border}`,
                  borderRadius: 8,
                  fontSize: 13,
                  fontFamily: "inherit",
                  background: C.white,
                }}
              >
                <option value="">Tous</option>
                {MOTIFS_PAIEMENT.map((m) => (
                  /* @__PURE__ */ <option key={m} value={m}>
                    {m}
                  </option>
                ))}
              </select>
            </div>
            <div>
              <label style={{ fontSize: 11, fontWeight: 600, color: "#555", display: "block", marginBottom: 4 }}>
                Groupe
              </label>
              <select
                value={filterFiliere}
                onChange={(e) => setFilterFiliere(e.target.value)}
                style={{
                  padding: "8px 11px",
                  border: `1px solid ${C.border}`,
                  borderRadius: 8,
                  fontSize: 13,
                  fontFamily: "inherit",
                  background: C.white,
                }}
              >
                <option value="">Tous les groupes</option>
                {getFilieres(data).map((f) => (
                  /* @__PURE__ */ <option key={f} value={f}>
                    {f}
                  </option>
                ))}
              </select>
            </div>
            <input
              placeholder="🔍 Étudiant, N° inscription, référence..."
              value={searchOps}
              onChange={(e) => setSearchOps(e.target.value)}
              style={{
                flex: 1,
                minWidth: 200,
                padding: "9px 14px",
                border: `1px solid ${C.border}`,
                borderRadius: 8,
                fontSize: 13,
                fontFamily: "inherit",
              }}
            />
            {(dateDebut || dateFin || filterMotif || filterFiliere || searchOps) && (
              /* @__PURE__ */ <Btn
                small={true}
                variant="light"
                onClick={() => {
                  setDateDebut("");
                  setDateFin("");
                  setFilterMotif("");
                  setFilterFiliere("");
                  setSearchOps("");
                }}
              >
                ✕ Réinitialiser
              </Btn>
            )}
          </div>
          <div
            style={{
              background: C.navy + "10",
              border: `1px solid ${C.navy}30`,
              borderRadius: 10,
              padding: "12px 16px",
              marginBottom: 14,
              display: "flex",
              justifyContent: "space-between",
              alignItems: "center",
              gap: 14,
              flexWrap: "wrap",
            }}
          >
            <span style={{ fontSize: 13, color: C.navy }}>
              {opsFiltered.length} opération(s) {dateDebut || dateFin ? `entre ${dateDebut || "…"} et ${dateFin || "…"}` : ""}
            </span>
            <div style={{ display: "flex", alignItems: "center", gap: 14 }}>
              <span style={{ fontSize: 16, fontWeight: 700, color: C.navy }}>Total: {formatMAD(totalOpsFiltered)}</span>
              <Btn
                small={true}
                onClick={() => exportOperationsToExcel(opsFiltered, user?.token, toast)}
                disabled={opsFiltered.length === 0}
              >
                📥 Exporter en Excel
              </Btn>
            </div>
          </div>
          <Card style={{ overflow: "hidden" }}>
            <DataTable
              cols={[
                { key: "date", label: "Date", render: (v) => formatDateFR(v) },
                { key: "etudiantPrenom", label: "Prénom" },
                { key: "etudiantNom", label: "Nom" },
                { key: "numInscription", label: "N° Inscription" },
                { key: "anneeUniversitaire", label: "Année univ." },
                {
                  key: "motif",
                  label: "Motif",
                  render: (v) => /* @__PURE__ */ <Badge color={C.orange}>{v || "—"}</Badge>,
                },
                {
                  key: "montant",
                  label: "Montant",
                  render: (v) => /* @__PURE__ */ <strong style={{ color: C.green }}>{formatMAD(v)}</strong>,
                },
                { key: "mode", label: "Mode" },
                { key: "reference", label: "Référence" },
                { key: "recu", label: "Reçu", render: (v) => /* @__PURE__ */ <Badge color={C.navy}>{v}</Badge> },
                { key: "agent", label: "Agent" },
                {
                  key: "etudiantIdRef",
                  label: "",
                  render: (_, row) => (
                    /* @__PURE__ */ <Btn
                      small={true}
                      variant="navy"
                      onClick={() => {
                        const e = eps.find((x) => x.id === row.etudiantIdRef);
                        if (e) setShowFiche(e);
                      }}
                    >
                      Voir fiche
                    </Btn>
                  ),
                },
              ]}
              rows={opsFiltered}
              emptyMsg="Aucune opération de paiement sur cette période"
            />
          </Card>
        </div>
      )}
      {tab === "etats" && (
        /* @__PURE__ */ <div>
          <div className="eco-grid2" style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16 }}>
            <Card style={{ padding: 20 }}>
              <h4 style={{ color: C.navy, marginBottom: 12 }}>📊 Par filière</h4>
              {[...new Set(eps.map((e) => e.filiere))].map((f) => {
                const fEps = eps.filter((e) => e.filiere === f);
                const fPaye = fEps.reduce((s, e) => s + e.totalPaye, 0);
                const fTotal = fEps.reduce((s, e) => s + e.totalAPayer, 0);
                return (
                  /* @__PURE__ */ <div
                    key={f}
                    style={{ marginBottom: 10, padding: "8px 0", borderBottom: `1px solid ${C.light}` }}
                  >
                    <div style={{ fontSize: 12, fontWeight: 600, color: C.navy, marginBottom: 4 }}>
                      {f.slice(0, 35)}...
                    </div>
                    <div style={{ display: "flex", justifyContent: "space-between", fontSize: 12 }}>
                      <span style={{ color: C.muted }}>{fEps.length} étudiants</span>
                      <span style={{ color: C.green, fontWeight: 600 }}>
                        {formatMAD(fPaye)} / {formatMAD(fTotal)}
                      </span>
                    </div>
                    <div style={{ background: "#eee", borderRadius: 4, height: 4, marginTop: 4 }}>
                      <div
                        style={{
                          width: `${fTotal > 0 ? Math.min(Math.round((fPaye / fTotal) * 100), 100) : 0}%`,
                          height: "100%",
                          background: C.green,
                          borderRadius: 4,
                        }}
                      />
                    </div>
                  </div>
                );
              })}
            </Card>
            <Card style={{ padding: 20 }}>
              <h4 style={{ color: C.navy, marginBottom: 12 }}>💳 Par mode de paiement</h4>
              {["Espèces", "Chèque", "Virement", "Carte bancaire", "Paiement en ligne"].map((mode) => {
                const total = eps.reduce(
                  (s, e) => s + e.versements.filter((v) => v.mode === mode).reduce((a, v) => a + v.montant, 0),
                  0
                );
                const count = eps.reduce((s, e) => s + e.versements.filter((v) => v.mode === mode).length, 0);
                if (!count) return null;
                return (
                  /* @__PURE__ */ <div
                    key={mode}
                    style={{
                      display: "flex",
                      justifyContent: "space-between",
                      padding: "8px 0",
                      borderBottom: `1px solid ${C.light}`,
                      fontSize: 13,
                    }}
                  >
                    <span style={{ color: "#555" }}>
                      {mode} ({count} opération{count > 1 ? "s" : ""})
                    </span>
                    <strong style={{ color: C.navy }}>{formatMAD(total)}</strong>
                  </div>
                );
              })}
              <div
                style={{
                  marginTop: 14,
                  padding: "10px",
                  background: C.navy + "10",
                  borderRadius: 8,
                  textAlign: "center",
                  fontWeight: 700,
                  color: C.navy,
                  fontSize: 14,
                }}
              >
                Total encaissé: {formatMAD(totalPaye)}
              </div>
            </Card>
          </div>
          <Card style={{ padding: 20, marginTop: 16 }}>
            <h4 style={{ color: C.red, marginBottom: 12 }}>📋 État des impayés</h4>
            <DataTable
              cols={[
                { key: "prenom", label: "Prénom" },
                { key: "nom", label: "Nom" },
                {
                  key: "filiere",
                  label: "Filière",
                  render: (v) => /* @__PURE__ */ <span style={{ fontSize: 11 }}>{v.slice(0, 25)}...</span>,
                },
                {
                  key: "totalAPayer",
                  label: "Total à payer",
                  render: (v) => /* @__PURE__ */ <span style={{ fontWeight: 600 }}>{formatMAD(v)}</span>,
                },
                {
                  key: "totalPaye",
                  label: "Payé",
                  render: (v) => (
                    /* @__PURE__ */ <span style={{ color: C.green, fontWeight: 600 }}>{formatMAD(v)}</span>
                  ),
                },
                {
                  key: "id",
                  label: "Reste",
                  render: (_, r) => (
                    /* @__PURE__ */ <span style={{ color: C.red, fontWeight: 600 }}>
                      {formatMAD(r.totalAPayer - r.totalPaye)}
                    </span>
                  ),
                },
              ]}
              rows={eps.filter((e) => e.totalPaye < e.totalAPayer)}
              emptyMsg="✅ Aucun impayé !"
            />
          </Card>
        </div>
      )}
      {tab === "archives" && role === "administrateur" && (
        /* @__PURE__ */ <div>
          <div
            style={{
              background: "#fdeeec",
              border: "1px solid #f3b0a8",
              borderRadius: 8,
              padding: "10px 14px",
              fontSize: 12,
              color: C.red,
              marginBottom: 16,
            }}
          >
            🗄️ Archives permanentes : chaque paiement ou versement supprimé manuellement est conservé ici indéfiniment.
            La suppression de données n'est jamais automatique dans ECOGEST.
          </div>
          <Card style={{ padding: 20, marginBottom: 16 }}>
            <h4 style={{ color: C.navy, marginBottom: 12 }}>
              💰 Comptes de paiement supprimés ({(data.paiementsArchives || []).length})
            </h4>
            <DataTable
              cols={[
                { key: "nomEtudiant", label: "Étudiant" },
                { key: "numInscription", label: "N° Inscription" },
                {
                  key: "filiere",
                  label: "Filière",
                  render: (v) => /* @__PURE__ */ <span style={{ fontSize: 11 }}>{(v || "").slice(0, 30)}</span>,
                },
                { key: "totalAPayer", label: "Total à payer", render: (v) => formatMAD(v || 0) },
                { key: "totalPaye", label: "Payé", render: (v) => formatMAD(v || 0) },
                { key: "dateSuppression", label: "Supprimé le", render: (v) => formatDateFR(v) || v },
                { key: "supprimePar", label: "Supprimé par" },
                {
                  key: "_restaurer",
                  label: "Restaurer",
                  render: (_, row) => (
                    /* @__PURE__ */ <Btn small={true} variant="navy" onClick={() => handleRestaurerPaiement(row)}>
                      ♻️ Restaurer
                    </Btn>
                  ),
                },
              ]}
              rows={data.paiementsArchives || []}
              emptyMsg="Aucun compte de paiement archivé"
            />
          </Card>
          <Card style={{ padding: 20 }}>
            <h4 style={{ color: C.navy, marginBottom: 12 }}>
              💳 Versements supprimés ({(data.versementsArchives || []).length})
            </h4>
            <DataTable
              cols={[
                { key: "nomEtudiant", label: "Étudiant" },
                { key: "numInscription", label: "N° Inscription" },
                { key: "date", label: "Date versement", render: (v) => formatDateFR(v) || v },
                {
                  key: "montant",
                  label: "Montant",
                  render: (v) => /* @__PURE__ */ <span style={{ fontWeight: 600 }}>{formatMAD(v || 0)}</span>,
                },
                { key: "mode", label: "Mode" },
                { key: "motif", label: "Motif" },
                { key: "recu", label: "Reçu" },
                { key: "dateSuppression", label: "Supprimé le", render: (v) => formatDateFR(v) || v },
                { key: "supprimePar", label: "Supprimé par" },
                {
                  key: "_restaurer",
                  label: "Restaurer",
                  render: (_, row) => (
                    /* @__PURE__ */ <Btn small={true} variant="navy" onClick={() => handleRestaurerVersement(row)}>
                      ♻️ Restaurer
                    </Btn>
                  ),
                },
              ]}
              rows={data.versementsArchives || []}
              emptyMsg="Aucun versement archivé"
            />
          </Card>
        </div>
      )}
      {showFiche && (
        /* @__PURE__ */ <FichePaiementEtudiant
          ep={showFiche}
          onClose={() => setShowFiche(null)}
          onSave={updateEtudiantPaiement}
          user={user}
          role={role}
          anneeActive={data.configuration?.anneeUniversitaireActive || ANNEES_UNIVERSITAIRES[0]}
        />
      )}
    </div>
  );
}
const NAV_BY_ROLE = {
  directrice: [
    "dashboard",
    "filieres",
    "visiteurs",
    "etudiants",
    "paiements",
    "pointage",
    "calendrier",
    "demandes",
    "presences",
    "notes",
    "cours",
    "emploi",
    "professeurs",
  ],
  administrateur: [
    "dashboard",
    "filieres",
    "visiteurs",
    "etudiants",
    "paiements",
    "pointage",
    "calendrier",
    "demandes",
    "presences",
    "notes",
    "cours",
    "emploi",
    "professeurs",
    "parametrage",
    "config",
  ],
  professeur: ["presences", "notes", "pointage"],
  etudiant: ["calendrier", "paiements", "demandes", "notes", "presences", "emploi"],
};
const NAV_SECTIONS = {
  administratif: ["filieres", "visiteurs", "etudiants", "paiements", "pointage", "calendrier", "demandes"],
  pedagogie: ["presences", "notes"],
};
const NAV_ITEMS = {
  dashboard: { label: "Tableau de bord", icon: "🏠" },
  paiements: { label: "Paiements", icon: "💰" },
  pointage: { label: "Pointage", icon: "🕛" },
  filieres: { label: "Groupes", icon: "🎯" },
  visiteurs: { label: "Visiteurs", icon: "👥" },
  etudiants: { label: "Étudiants", icon: "🎓" },
  professeurs: { label: "Professeurs", icon: "👨‍🏫" },
  notes: { label: "Évaluations", icon: "📊" },
  presences: { label: "Présences", icon: "✅" },
  cours: { label: "Modules", icon: "📚" },
  emploi: { label: "Emploi du temps", icon: "📅" },
  calendrier: { label: "Calendrier", icon: "🗓️" },
  demandes: { label: "Demandes de documents", icon: "📨" },
  parametrage: { label: "Paramétrage", icon: "🛠️" },
  config: { label: "Config. SharePoint", icon: "⚙️" },
};
const APP_VERSION = "v3.175.0";
const APP_DATE = "08 Août 2026";
const MSAL_CONFIG = {
  clientId: "0d52fee7-c6cf-4374-8063-395b2add8c91",
  tenantId: "72182bf9-410a-4e42-b9f0-4f740b3959e9", // Calculé dynamiquement à partir de l'URL réellement utilisée pour accéder à
  // l'app (au lieu d'une valeur figée) : fonctionne automatiquement quelle que
  // soit l'adresse (ancienne ou nouvelle), tant qu'elle est aussi déclarée
  // comme URI de redirection autorisée dans Azure AD.
  redirectUri: window.location.origin + window.location.pathname,
  siteUrl: "https://ecolesuperieurederabat.sharepoint.com/sites/ECOGEST",
  scopes: ["User.Read", "profile", "email", "openid", "Sites.ReadWrite.All", "Files.ReadWrite.All"],
};
const USER_ROLES = {
  "i.elmzioui@esrmi.ma": { role: "administrateur", nom: "Iliass ELMZIOUI", titre: "Administrateur de l'application" },
  "fz.zniber@esrmi.ma": { role: "directrice", nom: "Fatima Zahra ZNIBER", titre: "Directrice" },
  "k.bouabid@esrmi.ma": { role: "directrice", nom: "Kawtar BOUABID", titre: "Membre ESRMI" },
  "a.bahalla@esrmi.ma": { role: "directrice", nom: "Adila BAHALLA", titre: "Membre ESRMI" },
  "s.elhantaoui@esrmi.ma": { role: "directrice", nom: "Said EL HANTAOUI", titre: "Membre ESRMI" },
  "k.oubdi@esrmi.ma": { role: "directrice", nom: "K. OUBDI", titre: "Membre ESRMI" },
  "a.malti@esrmi.ma": { role: "professeur", nom: "Abdelhamid MALTI", titre: "Professeur" },
};
function generateCodeVerifier() {
  const array = new Uint8Array(32);
  crypto.getRandomValues(array);
  return btoa(String.fromCharCode(...array))
    .replace(/\+/g, "-")
    .replace(/\//g, "_")
    .replace(/=/g, "");
}
async function generateCodeChallenge(verifier) {
  const encoder = new TextEncoder();
  const data = encoder.encode(verifier);
  const digest = await crypto.subtle.digest("SHA-256", data);
  return btoa(String.fromCharCode(...new Uint8Array(digest)))
    .replace(/\+/g, "-")
    .replace(/\//g, "_")
    .replace(/=/g, "");
}
async function msalLogin() {
  const codeVerifier = generateCodeVerifier();
  const codeChallenge = await generateCodeChallenge(codeVerifier);
  sessionStorage.setItem("code_verifier", codeVerifier);
  const allScopes = MSAL_CONFIG.scopes;
  const authorizeUrl = `https://login.microsoftonline.com/${MSAL_CONFIG.tenantId}/oauth2/v2.0/authorize?client_id=${MSAL_CONFIG.clientId}&response_type=code&redirect_uri=${encodeURIComponent(MSAL_CONFIG.redirectUri)}&scope=${encodeURIComponent(allScopes.join(" "))}&response_mode=query&code_challenge=${codeChallenge}&code_challenge_method=S256&prompt=select_account`;
  window.location.href = authorizeUrl;
}
async function exchangeCodeForToken(code) {
  const codeVerifier = sessionStorage.getItem("code_verifier");
  const allScopes = MSAL_CONFIG.scopes.join(" ");
  const response = await fetch(`https://login.microsoftonline.com/${MSAL_CONFIG.tenantId}/oauth2/v2.0/token`, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({
      client_id: MSAL_CONFIG.clientId,
      grant_type: "authorization_code",
      code,
      redirect_uri: MSAL_CONFIG.redirectUri,
      code_verifier: codeVerifier,
      scope: allScopes,
    }),
  });
  const data = await response.json();
  if (data.error) console.error("Token error:", data.error_description);
  return data.access_token;
}
async function getUserInfo(token) {
  const res = await fetch("https://graph.microsoft.com/v1.0/me?$select=displayName,mail,userPrincipalName,jobTitle", {
    headers: { Authorization: `Bearer ${token}` },
  });
  return await res.json();
}
const LOGO =
  "data:image/png;base64,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";
const LOGO_JPEG_B64 =
  "/9j/4AAQSkZJRgABAQAAAQABAAD/2wBDAAMCAgICAgMCAgIDAwMDBAYEBAQEBAgGBgUGCQgKCgkICQkKDA8MCgsOCwkJDRENDg8QEBEQCgwSExIQEw8QEBD/2wBDAQMDAwQDBAgEBAgQCwkLEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBD/wAARCAD6APoDASIAAhEBAxEB/8QAHwAAAQUBAQEBAQEAAAAAAAAAAAECAwQFBgcICQoL/8QAtRAAAgEDAwIEAwUFBAQAAAF9AQIDAAQRBRIhMUEGE1FhByJxFDKBkaEII0KxwRVS0fAkM2JyggkKFhcYGRolJicoKSo0NTY3ODk6Q0RFRkdISUpTVFVWV1hZWmNkZWZnaGlqc3R1dnd4eXqDhIWGh4iJipKTlJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXGx8jJytLT1NXW19jZ2uHi4+Tl5ufo6erx8vP09fb3+Pn6/8QAHwEAAwEBAQEBAQEBAQAAAAAAAAECAwQFBgcICQoL/8QAtREAAgECBAQDBAcFBAQAAQJ3AAECAxEEBSExBhJBUQdhcRMiMoEIFEKRobHBCSMzUvAVYnLRChYkNOEl8RcYGRomJygpKjU2Nzg5OkNERUZHSElKU1RVVldYWVpjZGVmZ2hpanN0dXZ3eHl6goOEhYaHiImKkpOUlZaXmJmaoqOkpaanqKmqsrO0tba3uLm6wsPExcbHyMnK0tPU1dbX2Nna4uPk5ebn6Onq8vP09fb3+Pn6/9oADAMBAAIRAxEAPwD9U6KKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAoooo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Yhum/DHxHrV/8BNU8RW3jIznw+qeLydZ1CBVuItMjMX2tEmCiT7RuzkZZshs19L0UAeF/tK6T4m1bWvAcelx6y2jx3eoNqpsLfVJ4x/o37nzk010nI352knaD1HSn+IvhjZeJL74VXsFr4nNtHP9j1c2+qavYoLOPTbto/PieYSKPtPk8y5kJ2BmNe40UAcV48h8RSeLPh2+irfmyi8QTtqxtywjFt/Zd6F87HBTzjDjPG/Z3xXgvgbS/wBpnQbHw14YLa5qF0vgvV9S0zWNWlLJBqc1vB5On6luOWkhuSxjdwd0LFfvRsW+r6KAPkzTLfxwvh3Xrbwb4U+Iltqd38O9US+uNTudUS6GuC3Xy8i4Bt5rhpjJ5clo5AGf4SmHaD4E8TP8FdW8O+F5PFmieINWv/Dn2u80+y1yyvLYNqMBu50bUXdXdUMzyGP5SA3mBgRX1jRQB8y6KPj3f6zfN4r0rXba7tPiHoFvNJaPILK70+GyiW5u7cBsLaSvl2Q/ddmRslSTX+Eug+LrS98T3viCHxOPEE0viUW32mz14Nhry4NqVmlkNgQYfL8vy0HBQLgg19RUUAfJPhzwT8a/C/hjxPZ3cGvxXmofC2b+zP7MvL+5W51fyRu883Eztb3yOVEYjCiQSyHcWjCr0r6J4jl/aC1XVPEMHiH+z47/AEdtLY2WvTW5jFpB5mx7WQWaKJt+7zkPzbi+Vr6RooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooA//9k=";
function LoginScreen({ onLogin }) {
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");
  const handleMicrosoftLogin = async () => {
    setLoading(true);
    setError("");
    try {
      await msalLogin();
    } catch (err) {
      setError("Erreur de connexion. Réessayez.");
      setLoading(false);
    }
  };
  return (
    /* @__PURE__ */ <div
      className="eco-login-shell"
      style={{
        height: "100vh",
        overflow: "hidden",
        fontFamily:
          "-apple-system, BlinkMacSystemFont, 'Segoe UI Variable', 'Segoe UI', Roboto, Helvetica, Arial, sans-serif",
        display: "flex",
        background: `linear-gradient(135deg, ${C.navy} 0%, ${C.teal} 100%)`,
      }}
    >
      <div
        className="eco-login-left"
        style={{
          flex: 1,
          display: "flex",
          flexDirection: "column",
          alignItems: "center",
          justifyContent: "center",
          padding: "24px 40px",
          color: "white",
          minWidth: 0,
        }}
      >
        <div
          className="eco-login-sun"
          style={{
            position: "relative",
            width: 170,
            height: 170,
            marginBottom: 16,
            display: "flex",
            alignItems: "center",
            justifyContent: "center",
          }}
        >
          <div
            style={{
              position: "absolute",
              width: 320,
              height: 320,
              background:
                "repeating-conic-gradient(from 0deg, rgba(255,255,255,0.28) 0deg 2.2deg, transparent 2.2deg 16deg)",
              borderRadius: "50%",
              animation: "sunrays-cw 350s linear infinite",
              pointerEvents: "none",
            }}
          />
          <div
            style={{
              position: "absolute",
              width: 255,
              height: 255,
              background:
                "repeating-conic-gradient(from 12deg, rgba(255,183,77,0.55) 0deg 6deg, transparent 6deg 27deg)",
              borderRadius: "50%",
              animation: "sunrays-ccw 210s linear infinite",
              pointerEvents: "none",
            }}
          />
          <div
            style={{
              position: "absolute",
              width: 205,
              height: 205,
              borderRadius: "50%",
              background:
                "radial-gradient(circle, rgba(255,255,255,0.5) 0%, rgba(255,222,163,0.22) 45%, rgba(255,255,255,0) 72%)",
              animation: "sunpulse 4.5s ease-in-out infinite",
              pointerEvents: "none",
            }}
          />
          <img
            src={LOGO}
            alt="ESRMI"
            style={{ width: 150, position: "relative", zIndex: 1, filter: "drop-shadow(0 8px 20px rgba(0,0,0,0.4))" }}
          />
          <style>{`@keyframes sunrays-cw{to{transform:rotate(360deg)}}@keyframes sunrays-ccw{to{transform:rotate(-360deg)}}@keyframes sunpulse{0%,100%{opacity:0.75;transform:scale(1)}50%{opacity:1;transform:scale(1.09)}}`}</style>
        </div>
        <h1
          className="eco-login-title"
          style={{
            fontSize: 19,
            fontWeight: 800,
            margin: 0,
            textAlign: "center",
            whiteSpace: "nowrap",
            letterSpacing: "-0.01em",
            textShadow: "0 2px 8px rgba(0,0,0,0.35)",
          }}
        >
          École Supérieure de Rabat en Management et Ingénierie
        </h1>
        <div
          style={{
            width: 60,
            height: 3,
            background: C.orange,
            borderRadius: 2,
            margin: "12px 0 16px",
            boxShadow: "0 2px 8px rgba(229,107,45,0.5)",
          }}
        />
        <div
          style={{
            background: "rgba(255,255,255,0.14)",
            border: "1px solid rgba(255,255,255,0.28)",
            borderRadius: 16,
            padding: "14px 24px",
            maxWidth: 360,
            textAlign: "center",
            boxShadow: "0 8px 24px rgba(0,0,0,0.2)",
          }}
        >
          <h2
            style={{
              margin: "0 0 6px",
              fontSize: 19,
              fontWeight: 800,
              color: C.orange,
              textShadow: "0 1px 4px rgba(0,0,0,0.3)",
            }}
          >
            ECOGEST
          </h2>
          <p
            style={{
              margin: 0,
              fontSize: 15.5,
              fontWeight: 700,
              color: "#ffffff",
              lineHeight: 1.5,
              letterSpacing: "0.02em",
            }}
          >
            Apprendre · Innover · Réussir
          </p>
        </div>
        <div style={{ marginTop: 18, display: "flex", gap: 14 }}>
          {[
            ["👥", "300+", "Utilisateurs"],
            ["📚", "8", "Filières"],
            ["🎓", "Bac+3/5", "Niveaux"],
          ].map(([icon, val, lbl]) => (
            /* @__PURE__ */ <div
              key={lbl}
              style={{
                textAlign: "center",
                background: "rgba(255,255,255,0.14)",
                border: "1px solid rgba(255,255,255,0.25)",
                borderRadius: 11,
                padding: "9px 13px",
                minWidth: 78,
                boxShadow: "0 4px 12px rgba(0,0,0,0.15)",
              }}
            >
              <div
                style={{
                  width: 26,
                  height: 26,
                  borderRadius: "50%",
                  background: "rgba(255,255,255,0.22)",
                  display: "flex",
                  alignItems: "center",
                  justifyContent: "center",
                  fontSize: 13,
                  margin: "0 auto 5px",
                }}
              >
                {icon}
              </div>
              <div style={{ fontSize: 14.5, fontWeight: 800, color: "#ffffff" }}>{val}</div>
              <div style={{ fontSize: 9.5, color: "rgba(255,255,255,0.85)", fontWeight: 600 }}>{lbl}</div>
            </div>
          ))}
        </div>
      </div>
      <div
        className="eco-login-right"
        style={{
          width: 420,
          background: "white",
          display: "flex",
          flexDirection: "column",
          justifyContent: "center",
          padding: "24px 38px",
          boxShadow: "-24px 0 70px rgba(0,0,0,0.28)",
          position: "relative",
          overflow: "hidden",
        }}
      >
        <div
          style={{
            position: "absolute",
            top: 0,
            left: 0,
            right: 0,
            height: 5,
            background: `linear-gradient(90deg, ${C.orange} 0%, ${C.navy} 100%)`,
          }}
        />
        <div style={{ marginBottom: 16 }}>
          <div
            style={{
              fontSize: 10.5,
              fontWeight: 800,
              color: C.orange,
              textTransform: "uppercase",
              letterSpacing: "0.12em",
              marginBottom: 6,
            }}
          >
            🔒 Portail sécurisé
          </div>
          <h2 style={{ fontSize: 23, fontWeight: 800, color: C.navy, margin: "0 0 5px" }}>Connexion</h2>
          <p style={{ fontSize: 12.5, color: "#4a5568", margin: 0, fontWeight: 500 }}>
            Utilisez votre compte institutionnel @esrmi.ma
          </p>
        </div>
        <div
          style={{
            background: "#eaf4ff",
            border: "1px solid #a9d2f5",
            borderRadius: 10,
            padding: "10px 14px",
            fontSize: 12,
            color: "#1a5088",
            marginBottom: 14,
            lineHeight: 1.5,
            fontWeight: 500,
          }}
        >
          🔐 Connexion sécurisée via <strong>Microsoft 365</strong>
          <br />
          <span style={{ color: "#334155" }}>Accès réservé aux membres de l'ESRMI</span>
        </div>
        <div
          style={{
            fontSize: 10,
            fontWeight: 800,
            color: "#5b5f6b",
            textTransform: "uppercase",
            letterSpacing: "0.08em",
            marginBottom: 6,
          }}
        >
          Rôles disponibles
        </div>
        <div
          style={{
            marginBottom: 14,
            borderRadius: 11,
            border: `1px solid ${C.border}`,
            overflow: "hidden",
            boxShadow: "0 2px 8px rgba(0,0,0,0.05)",
          }}
        >
          {[
            { icon: "👩‍💼", label: "Directrice et équipe administrative", color: C.purple },
            { icon: "⚙️", label: "Administrateur de l'application", color: C.navy },
            { icon: "👨‍🏫", label: "Professeurs", color: C.green },
            { icon: "🎓", label: "Étudiants", color: C.orange },
          ].map((r, i) => (
            /* @__PURE__ */ <div
              key={i}
              style={{
                display: "flex",
                alignItems: "center",
                gap: 10,
                padding: "7px 13px",
                borderBottom: i < 3 ? `1px solid ${C.light}` : "none",
                background: i % 2 === 0 ? C.white : C.light,
              }}
            >
              <span
                style={{
                  width: 26,
                  height: 26,
                  borderRadius: 8,
                  background: r.color + "20",
                  display: "flex",
                  alignItems: "center",
                  justifyContent: "center",
                  fontSize: 13,
                  flexShrink: 0,
                }}
              >
                {r.icon}
              </span>
              <span style={{ fontSize: 12.5, fontWeight: 700, color: "#232733" }}>{r.label}</span>
            </div>
          ))}
        </div>
        {error && (
          /* @__PURE__ */ <div
            style={{
              background: "#fff5f5",
              border: "1px solid #feb2b2",
              borderRadius: 8,
              padding: "9px 13px",
              fontSize: 12,
              color: C.red,
              marginBottom: 12,
              fontWeight: 600,
            }}
          >
            ⚠️ {error}
          </div>
        )}
        <button
          onClick={handleMicrosoftLogin}
          disabled={loading}
          onMouseEnter={(e) => {
            if (!loading) {
              e.currentTarget.style.transform = "translateY(-2px)";
              e.currentTarget.style.boxShadow = "0 8px 20px rgba(26,47,94,0.45)";
            }
          }}
          onMouseLeave={(e) => {
            e.currentTarget.style.transform = "translateY(0)";
            e.currentTarget.style.boxShadow = "0 4px 14px rgba(26,47,94,0.3)";
          }}
          style={{
            width: "100%",
            padding: "13px 20px",
            borderRadius: 11,
            border: `2px solid ${loading ? C.border : C.navy}`,
            background: loading ? C.light : `linear-gradient(135deg, ${C.navy} 0%, ${C.teal} 100%)`,
            cursor: loading ? "not-allowed" : "pointer",
            display: "flex",
            alignItems: "center",
            justifyContent: "center",
            gap: 10,
            fontFamily: "inherit",
            fontSize: 13.5,
            fontWeight: 800,
            color: loading ? "#999" : C.white,
            boxShadow: "0 4px 14px rgba(26,47,94,0.3)",
            transition: "all 0.2s",
            marginBottom: 10,
          }}
        >
          {loading ? (
            /* @__PURE__ */ <span>⏳ Redirection vers Microsoft...</span>
          ) : (
            /* @__PURE__ */ <React.Fragment>
              <svg width="18" height="18" viewBox="0 0 21 21">
                <rect x="1" y="1" width="9" height="9" fill="#f25022" />
                <rect x="11" y="1" width="9" height="9" fill="#7fba00" />
                <rect x="1" y="11" width="9" height="9" fill="#00a4ef" />
                <rect x="11" y="11" width="9" height="9" fill="#ffb900" />
              </svg>
              Se connecter avec Microsoft 365
            </React.Fragment>
          )}
        </button>
        <div
          style={{
            textAlign: "center",
            fontSize: 10.5,
            color: "#5b5f6b",
            lineHeight: 1.5,
            marginTop: 8,
            fontWeight: 500,
          }}
        >
          Support : <strong style={{ color: C.navy }}>i.elmzioui@esrmi.ma</strong>
        </div>
        <div
          style={{
            textAlign: "center",
            marginTop: 8,
            padding: "6px 12px",
            background: C.navy + "10",
            borderRadius: 8,
            border: `1px solid ${C.navy}25`,
          }}
        >
          <span style={{ fontSize: 10.5, color: C.navy, fontWeight: 800 }}>ECOGEST {APP_VERSION}</span>
          <span style={{ fontSize: 9.5, color: "#5b5f6b", marginLeft: 6, fontWeight: 600 }}>— {APP_DATE}</span>
          <div style={{ fontSize: 9.5, color: "#5b5f6b", marginTop: 2, fontWeight: 500 }}>
            SharePoint 365 · Paiements · Filières · Visiteurs
          </div>
        </div>
      </div>
    </div>
  );
}
let _spSiteId = null;
async function getSiteIdCached(token) {
  if (_spSiteId) return _spSiteId;
  try {
    const res = await fetch(
      "https://graph.microsoft.com/v1.0/sites/ecolesuperieurederabat.sharepoint.com:/sites/ECOGEST",
      { headers: { Authorization: `Bearer ${token}` } }
    );
    const data = await res.json();
    if (!data.id) throw new Error("Site introuvable: " + JSON.stringify(data));
    _spSiteId = data.id;
    console.log("✅ Site SharePoint ECOGEST connecté");
    return data.id;
  } catch (e) {
    console.error("getSiteId error:", e.message);
    return null;
  }
}
async function getDriveIdCached(token) {
  if (window._spDriveId) return window._spDriveId;
  try {
    const siteId = await getSiteIdCached(token);
    if (!siteId) return null;
    const res = await fetch(`https://graph.microsoft.com/v1.0/sites/${siteId}/drives`, {
      headers: { Authorization: `Bearer ${token}` },
    });
    const data = await res.json();
    if (data.error || !data.value?.length) return null;
    const docDrive =
      data.value.find(
        (d) =>
          d.name === "Documents" ||
          d.name === "Documents partages" ||
          d.name === "Documents partagés" ||
          d.driveType === "documentLibrary"
      ) || data.value[0];
    window._spDriveId = docDrive.id;
    return docDrive.id;
  } catch (e) {
    console.error("getDriveIdCached error:", e.message);
    return null;
  }
}
function buildOperationsExcelXML(rows) {
  const esc = (s) =>
    String(s ?? "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;");
  const headerCells = [
    "Date",
    "Prénom",
    "Nom",
    "N° Inscription",
    "Année Universitaire",
    "Motif",
    "Montant (MAD)",
    "Mode",
    "Référence",
    "Reçu",
    "Agent",
  ]
    .map((h) => `<Cell><Data ss:Type="String">${esc(h)}</Data></Cell>`)
    .join("");
  const bodyRows = rows
    .map((v) => {
      const cells = [
        formatDateFR(v.date),
        v.etudiantPrenom,
        v.etudiantNom,
        v.numInscription,
        v.anneeUniversitaire || "",
        v.motif || "",
        v.montant,
        v.mode || "",
        v.reference || "",
        v.recu || "",
        v.agent || "",
      ]
        .map((val, i) =>
          i === 6
            ? `<Cell><Data ss:Type="Number">${Number(val) || 0}</Data></Cell>`
            : `<Cell><Data ss:Type="String">${esc(val)}</Data></Cell>`
        )
        .join("");
      return `<Row>${cells}</Row>`;
    })
    .join("");
  return `<?xml version="1.0"?>
<?mso-application progid="Excel.Sheet"?>
<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet" xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">
<Styles><Style ss:ID="hdr"><Font ss:Bold="1"/><Interior ss:Color="#1A2F5E" ss:Pattern="Solid"/><Font ss:Color="#FFFFFF" ss:Bold="1"/></Style></Styles>
<Worksheet ss:Name="Opérations"><Table>
<Row ss:StyleID="hdr">${headerCells}</Row>
${bodyRows}
</Table></Worksheet>
</Workbook>`;
}
async function exportOperationsToExcel(rows, token, toast) {
  const xml = buildOperationsExcelXML(rows);
  const filename = `Operations_Paiements_${/* @__PURE__ */ new Date().toISOString().slice(0, 10)}.xls`;
  const blob = new Blob([xml], { type: "application/vnd.ms-excel" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
  if (!token) {
    toast?.("⚠️ Non connecté à SharePoint — fichier téléchargé localement uniquement");
    return;
  }
  try {
    const driveId = await getDriveIdCached(token);
    if (!driveId) {
      toast?.("⚠️ Impossible d'accéder au répertoire partagé — fichier téléchargé localement uniquement");
      return;
    }
    const res = await fetch(
      `https://graph.microsoft.com/v1.0/drives/${driveId}/root:/ECOGEST_Exports/${encodeURIComponent(filename)}:/content`,
      {
        method: "PUT",
        headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/vnd.ms-excel" },
        body: xml,
      }
    );
    if (!res.ok) {
      toast?.("⚠️ Échec de l'enregistrement dans le répertoire partagé");
      return;
    }
    toast?.(`Export enregistré dans SharePoint (Documents partagés/ECOGEST_Exports/${filename}) ✓`);
  } catch (e) {
    console.error("exportOperationsToExcel error:", e.message);
    toast?.("⚠️ Échec de l'enregistrement dans le répertoire partagé");
  }
}
const _colCache = {};
const _rawColsCache = {};
async function getRawCols(listName, token, siteId) {
  if (_rawColsCache[listName]) return _rawColsCache[listName];
  try {
    const res = await fetch(`https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listName}/columns`, {
      headers: { Authorization: `Bearer ${token}` },
    });
    const data = await res.json();
    const cols = data.value || [];
    _rawColsCache[listName] = cols;
    return cols;
  } catch (e) {
    console.warn("getRawCols error:", e.message);
    return [];
  }
}
async function getColMap(listName, token, siteId) {
  if (_colCache[listName]) return _colCache[listName];
  const cols = await getRawCols(listName, token, siteId);
  const map = {};
  cols.forEach((c) => {
    map[c.displayName] = c.name;
    map[c.name] = c.name;
    map[c.displayName.toLowerCase()] = c.name;
    map[c.name.toLowerCase()] = c.name;
  });
  map["Title"] = "Title";
  _colCache[listName] = map;
  console.log(
    `📋 Colonnes ${listName}:`,
    Object.keys(map)
      .filter((k) => !k.includes("lower") && map[k].startsWith("field"))
      .map((k) => k + "->" + map[k])
  );
  return map;
}
async function getReverseColMap(listName, token, siteId) {
  const cols = await getRawCols(listName, token, siteId);
  const reverse = {};
  cols.forEach((c) => {
    reverse[c.name] = c.displayName;
  });
  reverse["Title"] = "Title";
  return reverse;
}
async function getColTypeMap(listName, token, siteId) {
  const cols = await getRawCols(listName, token, siteId);
  const types = {};
  cols.forEach((c) => {
    let type = "text";
    if (c.number || c.currency) type = "number";
    else if (c.boolean) type = "boolean";
    types[c.name] = type;
  });
  return types;
}
async function deleteFromSPOnline(list, spId, token) {
  try {
    if (!token || !spId) {
      console.error("deleteFromSP: token ou ID manquant");
      return false;
    }
    const siteId = await getSiteIdCached(token);
    if (!siteId) return false;
    const listName = list.startsWith("ECOGEST_") ? list : "ECOGEST_" + list;
    const res = await fetch(`https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listName}/items/${spId}`, {
      method: "DELETE",
      headers: { Authorization: `Bearer ${token}` },
    });
    if (res.status === 204) {
      console.log("✅ Supprimé de SharePoint:", listName, spId);
      return true;
    } else {
      console.error("SP delete error:", res.status);
      return false;
    }
  } catch (e) {
    console.error("SP delete exception:", e.message);
    return false;
  }
}
async function saveToSPOnline(list, fields, token, spId) {
  try {
    if (!token) {
      console.error("saveToSP: pas de token");
      return null;
    }
    const siteId = await getSiteIdCached(token);
    if (!siteId) {
      console.error("saveToSP: site introuvable");
      return null;
    }
    const listName = list.startsWith("ECOGEST_") ? list : "ECOGEST_" + list.charAt(0).toUpperCase() + list.slice(1);
    const base = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listName}/items`;
    const headers = { Authorization: `Bearer ${token}`, "Content-Type": "application/json" };
    const colMap = await getColMap(listName, token, siteId);
    const colTypes = await getColTypeMap(listName, token, siteId);
    const titleVal =
      fields.Title ||
      fields.Nom ||
      fields.Titre ||
      fields.Intitule ||
      fields.Code ||
      String(Object.values(fields)[0] || "ECOGEST");
    const safeFields = { Title: String(titleVal) };
    Object.keys(fields).forEach((k) => {
      if (k === "Title") return;
      const internalName = colMap[k] || colMap[k.toLowerCase()] || k;
      const type = colTypes[internalName];
      const raw = fields[k];
      let val;
      if (type === "number") {
        const n = parseFloat(raw);
        val = isNaN(n) ? 0 : n;
      } else if (type === "boolean") {
        val = !!raw && raw !== "false";
      } else {
        val = raw === null || raw === void 0 ? "" : String(raw);
      }
      safeFields[internalName !== "Title" ? internalName : k] = val;
    });
    console.log(
      "📋 ColMap keys:",
      Object.keys(colMap)
        .filter((k) => !k.toLowerCase().includes("lower"))
        .slice(0, 10)
    );
    console.log(
      "🔢 Types colonnes:",
      Object.keys(colTypes)
        .filter((k) => colTypes[k] !== "text")
        .map((k) => `${k}:${colTypes[k]}`)
    );
    console.log("💾 SP →", listName, Object.keys(safeFields));
    if (spId) {
      const res = await fetch(`${base}/${spId}/fields`, { method: "PATCH", headers, body: JSON.stringify(safeFields) });
      const d = await res.json();
      if (d.error) {
        console.warn(
          "SP PATCH failed, retrying without 'Statut' (colonne à choix limité, cause fréquente de rejet):",
          d.error.message
        );
        const { Statut, ...safeFieldsNoStatut } = safeFields;
        if (Statut !== void 0) {
          const res2 = await fetch(`${base}/${spId}/fields`, {
            method: "PATCH",
            headers,
            body: JSON.stringify(safeFieldsNoStatut),
          });
          const d2 = await res2.json();
          if (d2.error) {
            console.error("SP PATCH error (sans Statut):", d2.error.message);
            return null;
          }
          console.log("✅ Mis à jour (sans Statut):", listName);
          return { id: spId, ...d2 };
        }
        console.error("SP PATCH error:", d.error.message);
        return null;
      }
      console.log("✅ Mis à jour:", listName);
      return { id: spId, ...d };
    } else {
      const res = await fetch(base, { method: "POST", headers, body: JSON.stringify({ fields: safeFields }) });
      const d = await res.json();
      if (d.error) {
        console.warn("SP save failed with fields, trying Title only:", d.error.message);
        const res2 = await fetch(base, {
          method: "POST",
          headers,
          body: JSON.stringify({ fields: { Title: safeFields.Title } }),
        });
        const d2 = await res2.json();
        if (d2.error) {
          console.error("SP save error (Title only):", d2.error.message);
          return null;
        }
        console.log(
          "✅ Enregistré (Title only):",
          listName,
          d2.id,
          "— Cliquez 'Configurer les listes' pour activer toutes les colonnes"
        );
        return d2;
      }
      console.log("✅ Enregistré dans SharePoint:", listName, d.id);
      return d;
    }
  } catch (e) {
    console.error("SP save exception:", e.message);
    return null;
  }
} // \u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550
//  MODE HORS LIGNE : cache local des donn\xE9es + file d'attente de synchronisation
// \u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550
// Toute \xE9criture (saveToSP / deleteFromSP) passe d\xE9sormais par cette
// couche : en ligne, elle se comporte exactement comme avant (aucun appelant
// existant n'a besoin d'\xEAtre modifi\xE9). Hors ligne, elle met l'op\xE9ration en
// file d'attente dans le navigateur et renvoie IMM\xC9DIATEMENT un r\xE9sultat
// "optimiste" (avec un identifiant temporaire pour une cr\xE9ation), pour que
// tout le code appelant continue de fonctionner sans savoir qu'il est hors ligne.
const OFFLINE_QUEUE_KEY = "ecogest_offline_queue_v1";
const OFFLINE_DATA_CACHE_KEY = "ecogest_data_cache_v1";
function isBrowserOnline() {
  return typeof navigator !== "undefined" ? navigator.onLine !== false : true;
}
function loadOfflineQueue() {
  try {
    const raw = localStorage.getItem(OFFLINE_QUEUE_KEY);
    return raw ? JSON.parse(raw) : [];
  } catch (e) {
    return [];
  }
}
function persistOfflineQueue(queue) {
  try {
    localStorage.setItem(OFFLINE_QUEUE_KEY, JSON.stringify(queue));
  } catch (e) {
    console.error("Impossible d'enregistrer la file hors ligne:", e.message);
  }
  window.dispatchEvent(new CustomEvent("ecogest-queue-changed"));
}
function pushToOfflineQueue(item) {
  const queue = loadOfflineQueue();
  queue.push(item);
  persistOfflineQueue(queue);
}
function genOfflineOpId() {
  return `op-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
}
function saveDataCache(data) {
  try {
    localStorage.setItem(OFFLINE_DATA_CACHE_KEY, JSON.stringify({ data, savedAt: Date.now() }));
  } catch (e) {
    console.error("Impossible d'enregistrer le cache local:", e.message);
  }
}
function loadDataCache() {
  try {
    const raw = localStorage.getItem(OFFLINE_DATA_CACHE_KEY);
    if (!raw) return null;
    const parsed = JSON.parse(raw);
    return parsed && parsed.data ? parsed : null;
  } catch (e) {
    return null;
  }
}
async function saveToSP(list, fields, token, spId) {
  if (!isBrowserOnline()) {
    const tempId = spId || `offline-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
    pushToOfflineQueue({
      opId: genOfflineOpId(),
      type: "save",
      list,
      fields,
      spId: spId || null,
      tempId,
      timestamp: Date.now(),
    });
    console.log("📥 Hors ligne — mise en file d'attente:", list, tempId);
    return { id: tempId, _offline: true };
  }
  return saveToSPOnline(list, fields, token, spId);
}
async function deleteFromSP(list, spId, token) {
  if (!isBrowserOnline()) {
    if (typeof spId === "string" && spId.startsWith("offline-")) {
      const queue = loadOfflineQueue().filter((q) => q.tempId !== spId);
      persistOfflineQueue(queue);
      console.log("🗑️ Hors ligne — création en attente annulée:", list, spId);
      return true;
    }
    pushToOfflineQueue({ opId: genOfflineOpId(), type: "delete", list, spId, timestamp: Date.now() });
    console.log("📥 Hors ligne — suppression mise en file d'attente:", list, spId);
    return true;
  }
  return deleteFromSPOnline(list, spId, token);
}
function resolveSpId(spId, remap) {
  if (!spId) return null;
  return remap[spId] || spId;
}
async function flushOfflineQueue(token) {
  const queue = loadOfflineQueue();
  if (queue.length === 0) return { reussies: 0, echouees: 0, remap: {} };
  const remap = {};
  const restantes = [];
  let reussies = 0,
    echouees = 0;
  for (const item of queue) {
    try {
      if (item.type === "save") {
        const spIdReel = resolveSpId(item.spId, remap);
        const result = await saveToSPOnline(item.list, item.fields, token, spIdReel);
        if (result && result.id) {
          remap[item.tempId] = result.id;
          reussies++;
        } else {
          restantes.push(item);
          echouees++;
        }
      } else if (item.type === "delete") {
        const spIdReel = resolveSpId(item.spId, remap);
        const ok = await deleteFromSPOnline(item.list, spIdReel, token);
        if (ok) reussies++;
        else {
          restantes.push(item);
          echouees++;
        }
      }
    } catch (e) {
      restantes.push(item);
      echouees++;
    }
  }
  persistOfflineQueue(restantes);
  return { reussies, echouees, remap };
}
const SP_LISTS_SCHEMA = {
  ECOGEST_Etudiants: [
    "Nom",
    "Prenom",
    "CIN",
    "DateNaissance",
    "LieuNaissance",
    "Genre",
    "Email",
    "EmailPersonnel",
    "Telephone",
    "Adresse",
    "Ville",
    "Quartier",
    "Filiere",
    "Niveau",
    "Annee",
    "Promotion",
    "Statut",
    "NumInscription",
    "ResponsableFinancier",
    "TarifGlobalFormation",
    "ModeReglement",
    "NotesConvention",
    "InscritPar",
    "BacAnnee",
    "BacMention",
    "BacSerie",
    "Etablissement",
    "Observations",
  ],
  ECOGEST_Professeurs: [
    "Nom",
    "Prenom",
    "CIN",
    "Email",
    "Telephone",
    "Specialite",
    "Matiere",
    "Grade",
    "TypeContrat",
    "Statut",
  ],
  ECOGEST_Filieres: [
    "Code",
    "Intitule",
    "Niveau",
    "Duree",
    "Capacite",
    "Frais",
    "Statut",
    "Description",
    "Responsable",
    "Couleur",
    "Accreditation",
    "TypeSeances",
  ],
  ECOGEST_Visiteurs: [
    "Nom",
    "Prenom",
    "CIN",
    "DateNaissance",
    "LieuNaissance",
    "Genre",
    "Email",
    "Telephone",
    "Adresse",
    "Ville",
    "Quartier",
    "BacAnnee",
    "BacMention",
    "BacSerie",
    "Etablissement",
    "FiliereS",
    "DateVisite",
    "Source",
    "Statut",
    "Observations",
    "AnneeUniversitaire",
    "RecuPar",
  ],
  ECOGEST_Notes: [
    "EtudiantId",
    "NomEtudiant",
    "NumInscription",
    "Module",
    "Filiere",
    "Semestre",
    "TypeEvaluation",
    "Note",
    "DateEvaluation",
    "Enseignant",
    "EnseignantCIN",
  ],
  ECOGEST_Presences: [
    "EtudiantId",
    "NomEtudiant",
    "NumInscription",
    "Module",
    "CodeModule",
    "DatePresence",
    "Heure",
    "Statut",
    "Motif",
    "Enseignant",
    "EnseignantCIN",
    "Filiere",
    "Seance",
    "Semestre",
    "DocumentJustification",
    "JustifiePar",
    "DateJustification",
  ],
  ECOGEST_Cours: [
    "Intitule",
    "Departement",
    "ResponsablePedagogique",
    "Semestre",
    "Filiere",
    "Professeur",
    "ProfesseurCIN",
    "HeuresTotal",
    "Salle",
    "DateCC",
    "DateExamen",
  ],
  ECOGEST_EmploiDuTemps: [
    "Filiere",
    "Semestre",
    "JourSemaine",
    "Seance",
    "Module",
    "Professeur",
    "ProfesseurCIN",
    "Salle",
    "TypeRecurrence",
    "DateDebut",
    "DateFin",
    "DatesExclues",
    "EstRattrapage",
    "LienTeams",
    "LienEnregistrements",
  ],
  ECOGEST_Annonces: ["Titre", "Contenu", "Auteur", "Cible", "DateAnnonce", "DateExpiration", "Priorite", "Statut"],
  ECOGEST_Paiements: [
    "EtudiantId",
    "NumInscription",
    "NomEtudiant",
    "Filiere",
    "AnneeScolaire",
    "TotalAPayer",
    "TotalPaye",
    "Solde",
    "Statut",
    "RemisesJSON",
    "PenalitesJSON",
  ], // Chaque versement est un enregistrement isolé (une ligne = un paiement),
  // jamais un tableau JSON dans une seule cellule — c'est le correctif définitif
  // au problème de troncature/duplication rencontré avec l'ancien modèle.
  ECOGEST_Versements: [
    "EtudiantId",
    "NumInscription",
    "NomEtudiant",
    "Filiere",
    "AnneeUniversitaire",
    "DateVersement",
    "Montant",
    "Mode",
    "Reference",
    "Banque",
    "NumCheque",
    "Motif",
    "Recu",
    "Agent",
  ], // Archives permanentes : copie conserv\xE9e ind\xE9finiment de tout paiement/versement
  // supprim\xE9 manuellement par un administrateur. La suppression reste toujours
  // une action manuelle explicite \u2014 jamais automatique \u2014 mais rien n'est perdu.
  ECOGEST_PaiementsArchives: [
    "EtudiantId",
    "NumInscription",
    "NomEtudiant",
    "Filiere",
    "AnneeScolaire",
    "TotalAPayer",
    "TotalPaye",
    "Solde",
    "Statut",
    "RemisesJSON",
    "PenalitesJSON",
    "DateSuppression",
    "SupprimePar",
  ],
  ECOGEST_VersementsArchives: [
    "EtudiantId",
    "NumInscription",
    "NomEtudiant",
    "Filiere",
    "AnneeUniversitaire",
    "DateVersement",
    "Montant",
    "Mode",
    "Reference",
    "Banque",
    "NumCheque",
    "Motif",
    "Recu",
    "Agent",
    "DateSuppression",
    "SupprimePar",
  ],
  ECOGEST_PointageFiches: [
    "ProfesseurId",
    "ProfesseurNom",
    "Mois",
    "Statut",
    "DateCreation",
    "DateSoumission",
    "DateValidation",
    "ValideePar",
    "DateAccuseReception",
    "AccuseReceptionPar",
    "DateTraitement",
    "TraitePar",
  ],
  ECOGEST_PointageSeances: [
    "FicheId",
    "ProfesseurId",
    "ProfesseurNom",
    "Statut",
    "Confirmee",
    "DatePointage",
    "Module",
    "Filiere",
    "Semestre",
    "HeureDebut",
    "HeureFin",
    "NbHeures",
    "Commentaire",
    "ValideePar",
    "DateValidation",
  ], // Paramétrage général de l'application (un seul enregistrement) : année
  // universitaire active à appliquer automatiquement à toutes les opérations.
  ECOGEST_Parametres: ["AnneeUniversitaireActive", "RamadanDebut", "RamadanFin", "FilieresAvecTauxAbsence"],
  ECOGEST_SemestresDates: ["Semestre", "DateDebut", "DateFin"],
  ECOGEST_DemandesDocuments: [
    "EtudiantId",
    "NomEtudiant",
    "NumInscription",
    "Filiere",
    "TypeDocument",
    "DateDemande",
    "Statut",
    "Commentaire",
    "DateTraitement",
    "TraitePar",
  ],
  ECOGEST_Calendrier: ["TypeEvenement", "DateDebut", "DateFin", "Description", "Filiere"],
};
const SP_NUMBER_FIELDS = /* @__PURE__ */ new Set([
  "TotalAPayer",
  "TotalPaye",
  "Solde",
  "Montant",
  "TarifGlobalFormation",
  "NoteCC",
  "NoteExamen",
  "NoteFinal",
  "Coef",
]);
const SP_MULTILINE_FIELDS = /* @__PURE__ */ new Set([
  "RemisesJSON",
  "PenalitesJSON",
  "Echeances",
  "Observations",
  "NotesConvention",
  "LienTeams",
  "LienEnregistrements",
]);
function spColumnBodyFor(name) {
  if (SP_NUMBER_FIELDS.has(name)) return { number: { decimalPlaces: "automatic" } };
  if (SP_MULTILINE_FIELDS.has(name)) return { text: { allowMultipleLines: true } };
  return { text: {} };
}
async function setupSharePointLists(token) {
  const siteId = await getSiteIdCached(token);
  if (!siteId) return;
  const headers = { Authorization: `Bearer ${token}`, "Content-Type": "application/json" };
  const base = `https://graph.microsoft.com/v1.0/sites/${siteId}`;
  for (const [listName, columns] of Object.entries(SP_LISTS_SCHEMA)) {
    try {
      const checkRes = await fetch(`${base}/lists/${listName}`, { headers });
      const checkData = await checkRes.json();
      if (checkData.error && checkData.error.code === "itemNotFound") {
        const createRes = await fetch(`${base}/lists`, {
          method: "POST",
          headers,
          body: JSON.stringify({ displayName: listName, list: { template: "genericList" } }),
        });
        const listData = await createRes.json();
        if (listData.error) {
          console.error("Create list error:", listName, listData.error.message);
          continue;
        }
        console.log("✅ Liste créée:", listName);
        for (const col of columns) {
          await fetch(`${base}/lists/${listName}/columns`, {
            method: "POST",
            headers,
            body: JSON.stringify({ name: col, ...spColumnBodyFor(col) }),
          });
        }
        console.log("✅ Colonnes ajoutées:", listName, columns.length);
      } else {
        const colRes = await fetch(`${base}/lists/${listName}/columns`, { headers });
        const colData = await colRes.json();
        const existingColsData = colData.value || [];
        const existingCols = existingColsData.map((c) => c.displayName);
        for (const col of columns) {
          if (!existingCols.includes(col)) {
            try {
              const addRes = await fetch(`${base}/lists/${listName}/columns`, {
                method: "POST",
                headers,
                body: JSON.stringify({ name: col, ...spColumnBodyFor(col) }),
              });
              const addData = await addRes.json();
              if (addData.error) console.error("❌ Échec création colonne:", listName, col, addData.error.message);
              else console.log("✅ Colonne ajoutée:", listName, col);
            } catch (e) {
              console.error("❌ Erreur réseau création colonne:", listName, col, e.message);
            }
          } else {
            const existingCol = existingColsData.find((c) => c.displayName === col);
            const wrongNumber = SP_NUMBER_FIELDS.has(col) && existingCol?.text && !existingCol.number;
            const wrongMultiline =
              SP_MULTILINE_FIELDS.has(col) && existingCol?.text && !existingCol.text.allowMultipleLines;
            if (existingCol && (wrongNumber || wrongMultiline)) {
              try {
                const patchRes = await fetch(`${base}/lists/${listName}/columns/${existingCol.id}`, {
                  method: "PATCH",
                  headers,
                  body: JSON.stringify(spColumnBodyFor(col)),
                });
                const patchData = await patchRes.json();
                if (patchData.error)
                  console.warn(
                    "⚠️ Correction de type impossible (à faire manuellement dans SharePoint):",
                    listName,
                    col,
                    patchData.error.message
                  );
                else console.log("🔧 Type de colonne corrigé:", listName, col);
              } catch (e) {
                console.warn("Column type fix error:", listName, col, e.message);
              }
            }
          }
        }
      }
      delete _colCache[listName];
      delete _rawColsCache[listName];
    } catch (e) {
      console.error("Setup error:", listName, e.message);
    }
  }
  console.log("✅ Configuration SharePoint terminée");
}
async function loadAllFromSP(token) {
  console.log("📊 Chargement depuis SharePoint ECOGEST...");
  const siteId = await getSiteIdCached(token);
  if (!siteId) {
    console.warn("Site SharePoint inaccessible");
    return null;
  }
  const base = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists`;
  const headers = { Authorization: `Bearer ${token}` };
  const fetchList = async (name) => {
    try {
      const res = await fetch(`${base}/${name}/items?expand=fields&$top=500`, { headers });
      const d = await res.json();
      if (d.error) {
        console.warn(`Erreur liste ${name}:`, d.error.message);
        return [];
      }
      const reverseMap = await getReverseColMap(name, token, siteId);
      return (d.value || []).map((i) => {
        const fields = i.fields || {};
        Object.keys(fields).forEach((k) => {
          const disp = reverseMap[k];
          if (disp && disp !== k && fields[disp] === void 0) fields[disp] = fields[k];
        });
        if ((d.value || []).indexOf(i) === 0) console.log(`📋 ${name}:`, JSON.stringify(fields).slice(0, 500));
        return { _spId: i.id, ...fields };
      });
    } catch (e) {
      console.warn(`Erreur chargement ${name}:`, e.message);
      return [];
    }
  };
  const g = (item, ...keys) => {
    for (const k of keys)
      if (typeof k === "string" && item[k] !== void 0 && item[k] !== null && item[k] !== "") return item[k];
    return "";
  };
  const [
    etudiants,
    professeurs,
    filieres,
    visiteurs,
    notes,
    presences,
    cours,
    annonces,
    paiements,
    versementsSP,
    parametresSP,
    calendrier,
    emploiDuTemps,
    semestresDatesSP,
    demandesDocumentsSP,
    paiementsArchivesSP,
    versementsArchivesSP,
    pointageFichesSP,
    pointageSeancesSP,
  ] = await Promise.all([
    fetchList("ECOGEST_Etudiants"),
    fetchList("ECOGEST_Professeurs"),
    fetchList("ECOGEST_Filieres"),
    fetchList("ECOGEST_Visiteurs"),
    fetchList("ECOGEST_Notes"),
    fetchList("ECOGEST_Presences"),
    fetchList("ECOGEST_Cours"),
    fetchList("ECOGEST_Annonces"),
    fetchList("ECOGEST_Paiements"),
    fetchList("ECOGEST_Versements"),
    fetchList("ECOGEST_Parametres"),
    fetchList("ECOGEST_Calendrier"),
    fetchList("ECOGEST_EmploiDuTemps"),
    fetchList("ECOGEST_SemestresDates"),
    fetchList("ECOGEST_DemandesDocuments"),
    fetchList("ECOGEST_PaiementsArchives"),
    fetchList("ECOGEST_VersementsArchives"),
    fetchList("ECOGEST_PointageFiches"),
    fetchList("ECOGEST_PointageSeances"),
  ]);
  const versementsByEtudiant = {};
  versementsSP.forEach((v) => {
    const eid = g(v, "NumInscription");
    if (!versementsByEtudiant[eid]) versementsByEtudiant[eid] = [];
    versementsByEtudiant[eid].push({
      _spId: v._spId,
      id: v._spId,
      date: g(v, "DateVersement"),
      montant: parseFloat(g(v, "Montant") || 0),
      mode: g(v, "Mode"),
      reference: g(v, "Reference"),
      banque: g(v, "Banque"),
      numCheque: g(v, "NumCheque"),
      motif: g(v, "Motif"),
      anneeUniversitaire: g(v, "AnneeUniversitaire"),
      recu: g(v, "Recu"),
      agent: g(v, "Agent"),
    });
  });
  console.log(`✅ Chargé: ${etudiants.length} étudiants, ${filieres.length} filières, ${visiteurs.length} visiteurs`);
  return {
    etudiants: etudiants.map((e) => ({
      _spId: e._spId,
      id: e._spId,
      nom: e.Nom || e.Title?.split(" ").slice(1).join(" ") || "",
      prenom: e.Prenom || e.Title?.split(" ")[0] || "",
      cin: e.CIN || "",
      dateNaissance: e.DateNaissance || "",
      lieuNaissance: e.LieuNaissance || "",
      genre: e.Genre || "M",
      email: e.Email || "",
      emailPersonnel: e.EmailPersonnel || "",
      telephone: e.Telephone || "",
      adresse: e.Adresse || "",
      ville: e.Ville || "",
      quartier: e.Quartier || "",
      filiere: e.Filiere || "",
      niveau: e.Niveau || "",
      annee: e.Annee || "",
      promotion: e.Promotion || "",
      statut: e.Statut || "Actif",
      numInscription: e.NumInscription || "",
      responsableFinancier: e.ResponsableFinancier || "", // Nouveau champ unique "Tarif global de la formation". Si absent (étudiant
      // créé avant ce changement), on retombe sur la somme des anciens champs
      // par année pour ne perdre aucune donnée déjà saisie.
      tarifGlobalFormation:
        parseFloat(e.TarifGlobalFormation || 0) ||
        Math.round(
          (parseFloat(e.MontantAnnee1 || 0) +
            parseFloat(e.MontantAnnee2 || 0) +
            parseFloat(e.MontantAnnee3 || 0) +
            parseFloat(e.MontantAnnee4 || 0) +
            parseFloat(e.MontantAnnee5 || 0)) /
            500
        ) * 500,
      modeReglement: e.ModeReglement || "Trimestriel",
      notesConvention: e.NotesConvention || "",
      inscritPar: e.InscritPar || "",
      bacAnnee: e.BacAnnee || "",
      bacMention: e.BacMention || "Passable",
      bacSerie: e.BacSerie || "",
      etablissement: e.Etablissement || "",
      observations: e.Observations || "",
    })),
    professeurs: professeurs.map((p) => ({
      _spId: p._spId,
      id: p._spId,
      nom: p.Nom || p.Title?.split(" ").slice(1).join(" ") || "",
      prenom: p.Prenom || p.Title?.split(" ")[0] || "",
      cin: p.CIN || "",
      email: p.Email || "",
      telephone: p.Telephone || "",
      matiere: p.Specialite || p.Matiere || "",
      grade: p.Grade || "",
      type: p.TypeContrat || p.Type || "Permanent",
      statut: p.Statut || "Actif",
    })),
    filieres: filieres.map((f) => {
      return {
        _spId: f._spId,
        id: f._spId,
        code: f.Code || "",
        intitule: f.Intitule || f.Title || "",
        niveau: f.Niveau || "",
        duree: parseInt(f.Duree || 3),
        capacite: parseInt(f.Capacite || 30),
        frais: parseInt(f.Frais || 18e3),
        statut: f.Statut || "Active",
        description: f.Description || "",
        responsable: f.Responsable || "",
        couleur: f.Couleur || "#1a2f5e",
        accreditation: f.Accreditation || "", // Détermine quels créneaux horaires afficher pour ce groupe dans
        // l'Emploi du temps : "Soir" (uniquement 19h-21h) par défaut, ou
        // "Jour" (tous les créneaux sauf 19h-21h).
        typeSeances: f.TypeSeances || "Soir",
      };
    }),
    visiteurs: visiteurs.map((v) => ({
      _spId: v._spId,
      id: v._spId,
      nom: v.Nom || v.field_1 || v.Title?.split(" ").slice(1).join(" ") || "",
      prenom: v.Prenom || v.field_2 || v.Title?.split(" ")[0] || "",
      cin: v.CIN || v.field_3 || "",
      dateNaissance: v.DateNaissance || v.field_4 || "",
      lieuNaissance: v.LieuNaissance || v.field_5 || "",
      genre: v.Genre || v.field_6 || "M",
      email: v.Email || v.field_7 || "",
      telephone: v.Telephone || v.field_8 || "",
      adresse: v.Adresse || v.field_9 || "",
      ville: v.Ville || v.field_10 || "",
      quartier: v.Quartier || v.field_11 || "",
      bacAnnee: v.BacAnnee || v.field_12 || "",
      bacMention: v.BacMention || v.field_13 || "",
      bacSerie: v.BacSerie || v.field_14 || "",
      etablissement: v.Etablissement || v.field_15 || "",
      filiereS: v.FiliereS || v.field_16 || "",
      dateVisite: v.DateVisite || v.field_17 || "",
      source: v.Source || v.field_18 || "",
      statut: v.Statut || v.field_19 || "Visiteur",
      observations: v.Observations || v.field_20 || "",
      anneeUniversitaire: v.AnneeUniversitaire || "",
      recuPar: v.RecuPar || "",
      tarif: [],
      documents: [],
    })),
    notes: notes.map((n) => ({
      _spId: n._spId,
      id: n._spId,
      etudiantId: g(n, "EtudiantId"),
      etudiant: g(n, "NomEtudiant") || g(n, "Etudiant"),
      numInscription: g(n, "NumInscription"),
      module: g(n, "Module"),
      filiere: g(n, "Filiere"),
      semestre: g(n, "Semestre"),
      typeEvaluation: g(n, "TypeEvaluation") || "Contrôle continu",
      note: parseFloat(g(n, "Note") || 0),
      date: g(n, "DateEvaluation") || g(n, "DateNote"),
      enseignant: g(n, "Enseignant"),
      enseignantCin: g(n, "EnseignantCIN"),
    })),
    presences: presences.map((p) => ({
      _spId: p._spId,
      id: p._spId,
      etudiantId: g(p, "EtudiantId"),
      etudiant: g(p, "NomEtudiant") || g(p, "Etudiant"),
      numInscription: g(p, "NumInscription"),
      module: g(p, "Module"),
      codeModule: g(p, "CodeModule"),
      date: g(p, "DatePresence"),
      heure: g(p, "Heure"),
      statut: g(p, "Statut"),
      motif: g(p, "Motif"),
      enseignant: g(p, "Enseignant"),
      enseignantCin: g(p, "EnseignantCIN"),
      filiere: g(p, "Filiere"),
      seance: g(p, "Seance"),
      semestre: g(p, "Semestre"),
      documentJustification: g(p, "DocumentJustification"),
      justifiePar: g(p, "JustifiePar"),
      dateJustification: g(p, "DateJustification"),
    })),
    cours: cours.map((c) => {
      const filiereRaw = g(c, "Filiere");
      return {
        _spId: c._spId,
        id: c._spId,
        intitule: g(c, "Intitule"),
        departement: g(c, "Departement"),
        responsablePedagogique: g(c, "ResponsablePedagogique"),
        filiere: filiereRaw,
        filieresListe: filiereRaw
          ? filiereRaw
              .split("|")
              .map((s) => s.trim())
              .filter(Boolean)
          : [],
        semestre: g(c, "Semestre"),
        professeur: g(c, "Professeur") || g(c, "Enseignant"),
        professeurCin: g(c, "ProfesseurCIN"),
        heures: parseInt(g(c, "HeuresTotal") || g(c, "Heures") || 0),
        salle: g(c, "Salle"),
        dateCC: g(c, "DateCC"),
        dateExamen: g(c, "DateExamen"),
      };
    }),
    emploiDuTemps: emploiDuTemps.map((e) => ({
      _spId: e._spId,
      id: e._spId,
      filiere: g(e, "Filiere"),
      semestre: g(e, "Semestre"),
      jour: g(e, "JourSemaine"),
      seance: g(e, "Seance"),
      module: g(e, "Module"),
      professeur: g(e, "Professeur"),
      professeurCin: g(e, "ProfesseurCIN"),
      salle: g(e, "Salle"),
      typeRecurrence: g(e, "TypeRecurrence") || "Hebdomadaire",
      dateDebut: g(e, "DateDebut"),
      dateFin: g(e, "DateFin"),
      datesExclues: (g(e, "DatesExclues") || "")
        .split(",")
        .map((s) => s.trim())
        .filter(Boolean),
      estRattrapage: g(e, "EstRattrapage") === "Oui",
      lienTeams: g(e, "LienTeams"),
      lienEnregistrements: g(e, "LienEnregistrements"),
    })),
    semestresDates: semestresDatesSP.map((s) => ({
      _spId: s._spId,
      id: s._spId,
      semestre: g(s, "Semestre"),
      dateDebut: g(s, "DateDebut"),
      dateFin: g(s, "DateFin"),
    })),
    demandesDocuments: demandesDocumentsSP.map((d) => ({
      _spId: d._spId,
      id: d._spId,
      etudiantId: g(d, "EtudiantId"),
      etudiant: g(d, "NomEtudiant"),
      numInscription: g(d, "NumInscription"),
      filiere: g(d, "Filiere"),
      typeDocument: g(d, "TypeDocument"),
      dateDemande: g(d, "DateDemande"),
      statut: g(d, "Statut") || "En attente",
      commentaire: g(d, "Commentaire"),
      dateTraitement: g(d, "DateTraitement"),
      traitePar: g(d, "TraitePar"),
    })),
    annonces: annonces.map((a) => ({
      _spId: a._spId,
      id: a._spId,
      titre: g(a, "Titre"),
      contenu: g(a, "Contenu"),
      auteur: g(a, "Auteur"),
      cible: g(a, "Cible"),
      date: g(a, "DateAnnonce"),
      dateExpiration: g(a, "DateExpiration"),
      priorite: g(a, "Priorite") || "Normale",
      statut: g(a, "Statut") || "Active",
    })),
    paiements:
      paiements.length > 0
        ? {
            etudiants_paiement: paiements.map((p) => {
              const eid = g(p, "NumInscription");
              return {
                _spId: p._spId,
                id: p._spId,
                etudiantId: g(p, "EtudiantId"),
                numInscription: eid,
                nom: (g(p, "NomEtudiant") || "").split(" ").slice(1).join(" "),
                prenom: (g(p, "NomEtudiant") || "").split(" ")[0],
                filiere: g(p, "Filiere"),
                totalAPayer: parseFloat(g(p, "TotalAPayer") || 0),
                totalPaye: parseFloat(g(p, "TotalPaye") || 0),
                versements: versementsByEtudiant[eid] || [],
                remises: JSON.parse(g(p, "RemisesJSON") || g(p, "Remises") || "[]"),
                penalites: JSON.parse(g(p, "PenalitesJSON") || g(p, "Penalites") || "[]"),
                historique: [],
                statut: g(p, "Statut") || "Inscrit",
                promotion: g(p, "AnneeScolaire") || "2026-2027",
                responsableFinancier: "",
              };
            }),
            parametres: DEMO_PAIEMENT.parametres,
            parametres_penalites: DEMO_PAIEMENT.parametres_penalites,
            parametres_alertes: DEMO_PAIEMENT.parametres_alertes,
          }
        : null,
    paiementsArchives: paiementsArchivesSP.map((p) => ({
      _spId: p._spId,
      id: p._spId,
      numInscription: g(p, "NumInscription"),
      nomEtudiant: g(p, "NomEtudiant"),
      filiere: g(p, "Filiere"),
      promotion: g(p, "AnneeScolaire"),
      totalAPayer: parseFloat(g(p, "TotalAPayer") || 0),
      totalPaye: parseFloat(g(p, "TotalPaye") || 0),
      solde: parseFloat(g(p, "Solde") || 0),
      statut: g(p, "Statut"),
      dateSuppression: g(p, "DateSuppression"),
      supprimePar: g(p, "SupprimePar"),
    })),
    versementsArchives: versementsArchivesSP.map((v) => ({
      _spId: v._spId,
      id: v._spId,
      numInscription: g(v, "NumInscription"),
      nomEtudiant: g(v, "NomEtudiant"),
      filiere: g(v, "Filiere"),
      anneeUniversitaire: g(v, "AnneeUniversitaire"),
      date: g(v, "DateVersement"),
      montant: parseFloat(g(v, "Montant") || 0),
      mode: g(v, "Mode"),
      reference: g(v, "Reference"),
      motif: g(v, "Motif"),
      recu: g(v, "Recu"),
      agent: g(v, "Agent"),
      dateSuppression: g(v, "DateSuppression"),
      supprimePar: g(v, "SupprimePar"),
    })),
    pointageFiches: pointageFichesSP.map((f) => ({
      _spId: f._spId,
      id: f._spId,
      professeurId: g(f, "ProfesseurId"),
      professeurNom: g(f, "ProfesseurNom"),
      mois: g(f, "Mois"),
      statut: g(f, "Statut") || "Brouillon",
      dateCreation: g(f, "DateCreation"),
      dateSoumission: g(f, "DateSoumission"),
      dateValidation: g(f, "DateValidation"),
      valideePar: g(f, "ValideePar"),
      dateAccuseReception: g(f, "DateAccuseReception"),
      accuseReceptionPar: g(f, "AccuseReceptionPar"),
      dateTraitement: g(f, "DateTraitement"),
      traitePar: g(f, "TraitePar"),
    })),
    pointageSeances: pointageSeancesSP.map((s) => ({
      _spId: s._spId,
      id: s._spId,
      ficheId: g(s, "FicheId"),
      professeurId: g(s, "ProfesseurId"),
      professeurNom: g(s, "ProfesseurNom"),
      date: g(s, "DatePointage"),
      module: g(s, "Module"),
      filiere: g(s, "Filiere"),
      semestre: g(s, "Semestre"),
      heureDebut: g(s, "HeureDebut"),
      heureFin: g(s, "HeureFin"),
      nbHeures: parseFloat(g(s, "NbHeures") || 0),
      statut: g(s, "Statut") || "EnAttente",
      confirmee: g(s, "Confirmee") === "true" || g(s, "Confirmee") === true,
      commentaire: g(s, "Commentaire"),
      valideePar: g(s, "ValideePar"),
      dateValidation: g(s, "DateValidation"),
    })), // universitaire active, appliquée automatiquement à toutes les opérations
    // (versements...) sans que l'utilisateur ait à la sélectionner à chaque fois.
    configuration: {
      _spId: parametresSP[0]?._spId || null,
      anneeUniversitaireActive: g(parametresSP[0] || {}, "AnneeUniversitaireActive") || ANNEES_UNIVERSITAIRES[0],
      ramadanDebut: g(parametresSP[0] || {}, "RamadanDebut") || "",
      ramadanFin: g(parametresSP[0] || {}, "RamadanFin") || "",
      filieresAvecTauxAbsence: (g(parametresSP[0] || {}, "FilieresAvecTauxAbsence") || "")
        .split("|")
        .map((s) => s.trim())
        .filter(Boolean),
    },
    calendrier: calendrier.map((c) => ({
      _spId: c._spId,
      id: c._spId,
      titre: g(c, "Titre", "Title"),
      type: g(c, "TypeEvenement") || "Autre",
      dateDebut: g(c, "DateDebut"),
      dateFin: g(c, "DateFin"),
      description: g(c, "Description"),
      filiere: g(c, "Filiere"),
    })),
  };
}
export default function App() {
  const [roleAuth, setRole] = useState(null);
  const [userAuth, setUser] = useState(null);
  const [previewRole, setPreviewRole] = useState(null);
  const [previewUser, setPreviewUser] = useState(null);
  const [showTestMode, setShowTestMode] = useState(false);
  const [testModeSearch, setTestModeSearch] = useState("");
  const [active, setActive] = useState("dashboard");
  const [mobileMenuOpen, setMobileMenuOpen] = useState(false);
  const [pointageFiltreCible, setPointageFiltreCible] = useState(null);
  const [data, setData] = useState({
    etudiants: [],
    professeurs: [],
    filieres: [],
    visiteurs: [],
    notes: [],
    presences: [],
    cours: [],
    annonces: [],
    paiements: null,
    configuration: { _spId: null, anneeUniversitaireActive: ANNEES_UNIVERSITAIRES[0] },
    calendrier: [],
    emploiDuTemps: [],
    semestresDates: [],
    demandesDocuments: [],
  });
  const [toastMsg, setToastMsg] = useState("");
  const [checking, setChecking] = useState(true);
  const [spLoading, setSpLoading] = useState(false);
  const [isOnline, setIsOnline] = useState(isBrowserOnline());
  const [pendingSyncCount, setPendingSyncCount] = useState(loadOfflineQueue().length);
  const [syncingQueue, setSyncingQueue] = useState(false);
  const cacheLoadedRef = useRef(false);
  const toast = (msg) => {
    setToastMsg(msg);
    setTimeout(() => setToastMsg(""), 3e3);
  };
  useEffect(() => {
    const cached = loadDataCache();
    if (cached && cached.data) {
      setData(cached.data);
      cacheLoadedRef.current = true;
      console.log(
        "📦 Données locales chargées depuis le cache (mode hors ligne prêt) —",
        new Date(cached.savedAt).toLocaleString()
      );
    }
  }, []);
  useEffect(() => {
    if (!cacheLoadedRef.current && (!data.etudiants || data.etudiants.length === 0)) return;
    saveDataCache(data);
  }, [data]);
  const runSync = useCallback(
    (token) => {
      if (!token || syncingQueue) return;
      setSyncingQueue(true);
      flushOfflineQueue(token)
        .then((resultat) => {
          setPendingSyncCount(loadOfflineQueue().length);
          if (resultat.reussies > 0) {
            toast(`🔄 ${resultat.reussies} modification(s) synchronisée(s) avec SharePoint ✓`);
          }
          if (resultat.echouees > 0) {
            toast(
              `⚠️ ${resultat.echouees} modification(s) n'ont pas pu être synchronisées — nouvel essai automatique bientôt`
            );
          }
          return loadAllFromSP(token);
        })
        .then((spData) => {
          if (spData) setData(spData);
        })
        .catch((err) => console.warn("Synchronisation échouée:", err.message))
        .finally(() => setSyncingQueue(false));
    },
    [syncingQueue]
  );
  useEffect(() => {
    const handleOnline = () => {
      setIsOnline(true);
      if (userAuth?.token) runSync(userAuth.token);
    };
    const handleOffline = () => {
      setIsOnline(false);
      toast(
        "🔌 Connexion perdue — l'application continue de fonctionner, vos modifications seront synchronisées automatiquement au retour d'internet"
      );
    };
    const handleQueueChanged = () => setPendingSyncCount(loadOfflineQueue().length);
    window.addEventListener("online", handleOnline);
    window.addEventListener("offline", handleOffline);
    window.addEventListener("ecogest-queue-changed", handleQueueChanged);
    return () => {
      window.removeEventListener("online", handleOnline);
      window.removeEventListener("offline", handleOffline);
      window.removeEventListener("ecogest-queue-changed", handleQueueChanged);
    };
  }, [userAuth, runSync]);
  // Filet de sécurité : si l'app démarre (ou se recharge) déjà connectée à
  // internet alors que des modifications hors ligne sont encore en attente
  // (ex : l'événement navigateur "online" n'a pas pu être capté pendant une
  // session précédente), on relance automatiquement la synchronisation dès
  // qu'un token SharePoint valide est disponible, sans attendre un nouveau
  // cycle hors-ligne/en-ligne.
  useEffect(() => {
    if (!userAuth?.token || !isBrowserOnline()) return;
    if (loadOfflineQueue().length === 0) return;
    runSync(userAuth.token);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [userAuth?.token]);
  useEffect(() => {
    if (roleAuth && active && NAV_ITEMS[active]) {
      document.title = `ECOGEST ${APP_VERSION} — ${NAV_ITEMS[active].label}`;
    }
  }, [active, roleAuth]);
  useEffect(() => {
    const orphans = (data.visiteurs || []).filter(
      (v) => v.statut === "Inscrit" && v.cin && !(data.etudiants || []).some((e) => e.cin === v.cin)
    );
    if (orphans.length === 0) return;
    setData((d) => ({
      ...d,
      visiteurs: (d.visiteurs || []).map((v) =>
        orphans.some((o) => o.id === v.id) ? { ...v, statut: "Visiteur" } : v
      ),
    }));
    if (userAuth?.token) {
      orphans.forEach((o) => {
        if (o._spId) saveToSP("ECOGEST_Visiteurs", { Statut: "Visiteur" }, userAuth.token, o._spId);
      });
    }
    toast(`${orphans.length} visiteur(s) sans étudiant correspondant repassé(s) en statut "Visiteur" ✓`);
  }, [data.visiteurs, data.etudiants]);
  useEffect(() => {
    if (!data.paiements?.etudiants_paiement || !userAuth?.token) return;
    const orphelins = data.paiements.etudiants_paiement.filter((p) => {
      if (!p.numInscription) return true;
      const et = (data.etudiants || []).find((e) => e.numInscription === p.numInscription);
      if (!et) return true;
      const memePersonne =
        (p.nom || "").trim().toLowerCase() === (et.nom || "").trim().toLowerCase() &&
        (p.prenom || "").trim().toLowerCase() === (et.prenom || "").trim().toLowerCase();
      return !memePersonne;
    });
    if (orphelins.length === 0) return; // IMPORTANT : la suppression de donn\xE9es financi\xE8res ne doit JAMAIS \xEAtre automatique.
    // On se contente ici de signaler le probl\xE8me dans la console pour investigation
    // manuelle par un administrateur \u2014 aucune suppression locale ni SharePoint n'est
    // d\xE9clench\xE9e automatiquement.
    console.warn(
      `⚠️ ${orphelins.length} compte(s) de paiement sans correspondance exacte avec un étudiant actuel (vérification manuelle recommandée, aucune suppression automatique effectuée) :`,
      orphelins.map((p) => `${p.prenom} ${p.nom} (${p.numInscription})`)
    );
  }, [data.etudiants, data.paiements, userAuth]);
  useEffect(() => {
    if (!userAuth?.token) return;
    const intervalId = setInterval(
      () => {
        if (!isBrowserOnline()) return;
        loadAllFromSP(userAuth.token)
          .then((spData) => {
            if (spData) setData(spData);
          })
          .catch((err) => console.warn("Rafraîchissement automatique échoué:", err.message));
      },
      5 * 60 * 1e3
    );
    return () => clearInterval(intervalId);
  }, [userAuth]);
  const lastPageSyncRef = useRef(0);
  const isFirstActiveRef = useRef(true);
  const SYNC_MIN_INTERVAL_MS = 15 * 1e3;
  useEffect(() => {
    if (isFirstActiveRef.current) {
      isFirstActiveRef.current = false;
      return;
    }
    if (!userAuth?.token || !isBrowserOnline()) return;
    const now = Date.now();
    if (now - lastPageSyncRef.current < SYNC_MIN_INTERVAL_MS) return;
    lastPageSyncRef.current = now;
    setSpLoading(true);
    loadAllFromSP(userAuth.token)
      .then((spData) => {
        if (spData) setData(spData);
      })
      .catch((err) => console.warn("Synchronisation à la navigation échouée:", err.message))
      .finally(() => setSpLoading(false));
  }, [active]);
  useEffect(() => {
    const urlParams = new URLSearchParams(window.location.search);
    const code = urlParams.get("code");
    const error = urlParams.get("error");
    if (error) {
      console.error("Auth error:", error, urlParams.get("error_description"));
      window.history.replaceState({}, document.title, window.location.pathname);
      setChecking(false);
      return;
    }
    if (code) {
      window.history.replaceState({}, document.title, window.location.pathname);
      exchangeCodeForToken(code)
        .then((token) => {
          if (!token) {
            setChecking(false);
            return;
          }
          return getUserInfo(token).then((userInfo) => {
            const email = (userInfo.mail || userInfo.userPrincipalName || "").toLowerCase();
            const userConfig = USER_ROLES[email];
            if (userConfig) {
              const userObj = { email, nom: userConfig.nom, titre: userConfig.titre, token };
              setUser(userObj);
              setRole(userConfig.role);
              setActive(NAV_BY_ROLE[userConfig.role][0]);
              setChecking(false);
              try {
                const parts = token.split(".");
                const payload = JSON.parse(atob(parts[1].replace(/-/g, "+").replace(/_/g, "/")));
                console.log("Token scopes:", payload.scp || payload.roles || "none");
                console.log("Token audience:", payload.aud);
                console.log("Token expires:", new Date(payload.exp * 1e3).toLocaleString());
              } catch (e) {
                console.log("Token decode error:", e);
              }
              lastPageSyncRef.current = Date.now();
              setSpLoading(true);
              setupSharePointLists(token)
                .then(() => loadAllFromSP(token))
                .then((spData) => {
                  if (spData) setData(spData);
                  setSpLoading(false);
                })
                .catch((err) => {
                  console.warn("SP load failed:", err);
                  setSpLoading(false);
                });
            } else {
              lastPageSyncRef.current = Date.now();
              setSpLoading(true);
              setupSharePointLists(token)
                .then(() => loadAllFromSP(token))
                .then((spData) => {
                  const matched = (spData?.etudiants || []).find(
                    (e) => (e.email || "").toLowerCase() === email || (e.emailPersonnel || "").toLowerCase() === email
                  );
                  const matchedProf = !matched
                    ? (spData?.professeurs || []).find((p) => (p.email || "").toLowerCase() === email)
                    : null;
                  if (matched) {
                    const userObj = {
                      email,
                      nom: `${matched.prenom} ${matched.nom}`,
                      titre: "Étudiant",
                      token,
                      etudiantId: matched.id,
                      numInscription: matched.numInscription,
                    };
                    setUser(userObj);
                    setRole("etudiant");
                    setActive(NAV_BY_ROLE.etudiant[0]);
                    if (spData) setData(spData);
                  } else if (matchedProf) {
                    const userObj = {
                      email,
                      nom: `${matchedProf.prenom} ${matchedProf.nom}`,
                      titre: "Professeur",
                      token,
                      professeurId: matchedProf.id,
                    };
                    setUser(userObj);
                    setRole("professeur");
                    setActive(NAV_BY_ROLE.professeur[0]);
                    if (spData) setData(spData);
                  } else {
                    alert(`Accès non autorisé pour ${email}.
Contactez l'administrateur : i.elmzioui@esrmi.ma`);
                  }
                  setSpLoading(false);
                  setChecking(false);
                })
                .catch((err) => {
                  console.warn("SP load failed:", err);
                  alert(`Accès non autorisé pour ${email}.
Contactez l'administrateur : i.elmzioui@esrmi.ma`);
                  setSpLoading(false);
                  setChecking(false);
                });
            }
          });
        })
        .catch((err) => {
          console.error(err);
          setChecking(false);
        });
    } else {
      setChecking(false);
    }
  }, []);
  const handleLogin = (r, userInfo) => {
    setRole(r);
    setUser(userInfo);
    setActive(NAV_BY_ROLE[r][0]);
  };
  if (checking)
    return (
      /* @__PURE__ */ <div
        style={{
          minHeight: "100vh",
          background: `linear-gradient(135deg, ${C.navy} 0%, #2d4a8a 100%)`,
          display: "flex",
          alignItems: "center",
          justifyContent: "center",
          fontFamily:
            "-apple-system, BlinkMacSystemFont, 'Segoe UI Variable', 'Segoe UI', Roboto, Helvetica, Arial, sans-serif",
        }}
      >
        <div style={{ color: "white", textAlign: "center" }}>
          <div style={{ fontSize: 40, marginBottom: 16 }}>⏳</div>
          <div style={{ fontSize: 16, fontWeight: 600 }}>Vérification de votre compte Microsoft...</div>
        </div>
      </div>
    );
  if (!roleAuth) return /* @__PURE__ */ <LoginScreen onLogin={handleLogin} />;
  const role = previewRole || roleAuth;
  const user = previewRole ? previewUser : userAuth;
  const navItems = NAV_BY_ROLE[role];
  const PAGES = {
    dashboard: Dashboard,
    paiements: Paiements,
    pointage: Pointage,
    filieres: Filieres,
    visiteurs: Visiteurs,
    etudiants: Etudiants,
    professeurs: Professeurs,
    notes: Notes,
    presences: Presences,
    cours: Cours,
    emploi: EmploiDuTemps,
    calendrier: Calendrier,
    demandes: DemandesDocuments,
    config: Configuration,
    parametrage: Parametrage,
  };
  const Page = PAGES[active];
  const roleInfo = ROLES[role];
  return (
    /* @__PURE__ */ <div
      className="eco-app-shell"
      style={{
        display: "flex",
        height: "100vh",
        fontFamily:
          "-apple-system, BlinkMacSystemFont, 'Segoe UI Variable', 'Segoe UI', Roboto, Helvetica, Arial, sans-serif",
        background: C.light,
        overflow: "hidden",
      }}
    >
      <button
        className="eco-hamburger"
        onClick={() => setMobileMenuOpen((o) => !o)}
        aria-label="Menu"
        style={{
          display: "none",
          position: "fixed",
          top: 12,
          left: 12,
          zIndex: 1100,
          width: 40,
          height: 40,
          borderRadius: 10,
          border: `1px solid ${C.border}`,
          background: C.white,
          boxShadow: C.shadowSm,
          fontSize: 18,
          cursor: "pointer",
        }}
      >
        {mobileMenuOpen ? "✕" : "☰"}
      </button>
      {mobileMenuOpen && (
        /* @__PURE__ */ <div
          className="eco-sidebar-backdrop"
          onClick={() => setMobileMenuOpen(false)}
          style={{ display: "none", position: "fixed", inset: 0, background: "rgba(10,16,35,0.5)", zIndex: 1050 }}
        />
      )}
      <div
        className={"eco-sidebar" + (mobileMenuOpen ? " eco-sidebar-open" : "")}
        style={{
          width: 232,
          background: "#ffffff",
          display: "flex",
          flexDirection: "column",
          flexShrink: 0,
          overflowY: "auto",
          borderRight: `1px solid ${C.border}`,
          boxShadow: "1px 0 0 rgba(16,24,50,0.02)",
        }}
      >
        <div style={{ padding: "20px 18px 16px", borderBottom: `1px solid ${C.border}` }}>
          <div
            style={{
              fontSize: 10,
              color: C.muted,
              letterSpacing: "0.1em",
              textTransform: "uppercase",
              marginBottom: 4,
              fontWeight: 600,
            }}
          >
            École Supérieure de Rabat
          </div>
          <div style={{ fontSize: 20, fontWeight: 800, color: C.navy, letterSpacing: "-0.02em" }}>
            ECO<span style={{ color: C.orange }}>GEST</span>
          </div>
          <div style={{ fontSize: 10, color: "#aab0bd", marginTop: 3 }}>ECOGEST · SharePoint 365</div>
        </div>
        <div style={{ padding: "12px 12px 6px" }}>
          <div
            style={{
              background: C.light,
              borderRadius: 11,
              padding: "9px 11px",
              display: "flex",
              alignItems: "center",
              gap: 9,
              border: `1px solid ${C.border}`,
            }}
          >
            <div
              style={{
                width: 32,
                height: 32,
                borderRadius: "50%",
                background: roleInfo.color,
                display: "flex",
                alignItems: "center",
                justifyContent: "center",
                fontSize: 14,
                flexShrink: 0,
                boxShadow: C.shadowSm,
              }}
            >
              {roleInfo.icon}
            </div>
            <div style={{ overflow: "hidden" }}>
              <div
                style={{
                  fontSize: 11,
                  fontWeight: 700,
                  color: C.navy,
                  whiteSpace: "nowrap",
                  overflow: "hidden",
                  textOverflow: "ellipsis",
                }}
              >
                {user?.nom || roleInfo.label}
              </div>
              <div style={{ fontSize: 10, color: roleInfo.color, fontWeight: 600 }}>
                {user?.titre || roleInfo.label}
              </div>
              <div
                style={{
                  fontSize: 9,
                  color: "#aab0bd",
                  whiteSpace: "nowrap",
                  overflow: "hidden",
                  textOverflow: "ellipsis",
                }}
              >
                {user?.email || ""}
              </div>
            </div>
          </div>
        </div>
        <nav style={{ flex: 1, padding: "10px 10px", display: "flex", flexDirection: "column", gap: 5 }}>
          {(() => {
            const SECTION_META = {
              administratif: {
                label: "Administratif",
                icon: "🗂️",
                gradient: "linear-gradient(135deg, #a8460f 0%, #e56b2d 55%, #f4a45f 100%)",
              },
              pedagogie: {
                label: "Pédagogie",
                icon: "📘",
                gradient: "linear-gradient(135deg, #1f222a 0%, #454b58 45%, #8b93a3 100%)",
              },
              parametrage: {
                label: "Paramétrage",
                icon: "⚙️",
                gradient: "linear-gradient(135deg, #0c1a38 0%, #1a2f5e 40%, #5a83d6 100%)",
              },
            };
            const sectionOf = (key) =>
              NAV_SECTIONS.administratif.includes(key)
                ? "administratif"
                : NAV_SECTIONS.pedagogie.includes(key)
                  ? "pedagogie"
                  : ["parametrage", "config", "cours", "emploi", "professeurs"].includes(key)
                    ? "parametrage"
                    : null;
            const renderBtn = (key, onDark) => {
              const item = NAV_ITEMS[key];
              const LABELS_ETUDIANT = {
                paiements: "Mes paiements",
                notes: "Mes notes",
                presences: "Mes Absences",
                demandes: "Mes demandes",
              };
              const label = role === "etudiant" && LABELS_ETUDIANT[key] ? LABELS_ETUDIANT[key] : item.label;
              const isActive = active === key;
              const activeBg = onDark ? C.white : C.orange;
              return (
                /* @__PURE__ */ <button
                  key={key}
                  onClick={() => {
                    setActive(key);
                    setMobileMenuOpen(false);
                  }}
                  style={{
                    display: "flex",
                    alignItems: "center",
                    gap: 10,
                    width: "100%",
                    padding: "8px 10px",
                    border: "none",
                    cursor: "pointer",
                    textAlign: "left",
                    background: isActive ? activeBg : "transparent",
                    color: isActive ? (onDark ? C.navy : C.white) : onDark ? "#ffffff" : "#3d4356",
                    fontSize: 13,
                    fontFamily: "inherit",
                    fontWeight: isActive ? 700 : onDark ? 600 : 500,
                    borderRadius: 9,
                    marginBottom: 2,
                    boxShadow: isActive
                      ? onDark
                        ? "0 4px 12px rgba(0,0,0,0.3)"
                        : "0 4px 12px rgba(229,107,45,0.4)"
                      : "none",
                    textShadow: onDark && !isActive ? "0 1px 2px rgba(0,0,0,0.35)" : "none",
                    transition: "background 0.13s, color 0.13s, box-shadow 0.13s",
                  }}
                  onMouseEnter={(e) => {
                    if (!isActive) e.currentTarget.style.background = onDark ? "rgba(255,255,255,0.16)" : C.light;
                  }}
                  onMouseLeave={(e) => {
                    if (!isActive) e.currentTarget.style.background = "transparent";
                  }}
                >
                  <span
                    style={{
                      fontSize: 15,
                      width: 26,
                      height: 26,
                      borderRadius: 7,
                      flexShrink: 0,
                      display: "flex",
                      alignItems: "center",
                      justifyContent: "center",
                      background: onDark && !isActive ? "rgba(255,255,255,0.28)" : "transparent",
                      border: onDark && !isActive ? "1px solid rgba(255,255,255,0.25)" : "none",
                      boxShadow: onDark && !isActive ? "0 1px 3px rgba(0,0,0,0.2)" : "none",
                    }}
                  >
                    {item.icon}
                  </span>
                  {label}
                </button>
              );
            };
            const groups = [];
            navItems.forEach((key) => {
              const sec = sectionOf(key);
              const last = groups[groups.length - 1];
              if (last && last.sec === sec) last.items.push(key);
              else groups.push({ sec, items: [key] });
            });
            return groups.map((g, gi) => {
              if (!g.sec) return /* @__PURE__ */ <div key={gi}>{g.items.map((key) => renderBtn(key, false))}</div>;
              const meta = SECTION_META[g.sec];
              return (
                /* @__PURE__ */ <div
                  key={gi}
                  style={{
                    background: meta.gradient,
                    borderRadius: 12,
                    padding: "8px 6px",
                    boxShadow: "0 3px 10px rgba(0,0,0,0.18)",
                  }}
                >
                  <div style={{ display: "flex", alignItems: "center", gap: 6, padding: "2px 7px 6px" }}>
                    <span
                      style={{
                        fontSize: 12,
                        background: "rgba(255,255,255,0.28)",
                        borderRadius: 6,
                        width: 20,
                        height: 20,
                        display: "flex",
                        alignItems: "center",
                        justifyContent: "center",
                      }}
                    >
                      {meta.icon}
                    </span>
                    <span
                      style={{
                        fontSize: 10,
                        fontWeight: 800,
                        color: "#ffffff",
                        textShadow: "0 1px 2px rgba(0,0,0,0.3)",
                        textTransform: "uppercase",
                        letterSpacing: "0.09em",
                        textDecoration: "underline",
                        textUnderlineOffset: "3px",
                      }}
                    >
                      {meta.label}
                    </span>
                  </div>
                  {g.items.map((key) => renderBtn(key, true))}
                </div>
              );
            });
          })()}
        </nav>
        <div style={{ padding: "14px 18px", borderTop: `1px solid ${C.border}` }}>
          <div style={{ fontSize: 10, color: "#aab0bd", lineHeight: 1.6 }}>
            <span style={{ color: C.orange }}>● </span>SharePoint ECOGEST
            <br />
            ecolesuperieurederabat
            <br />
            .sharepoint.com/sites/ECOGEST
          </div>
          {roleAuth === "administrateur" && !previewRole && (
            /* @__PURE__ */ <button
              onClick={() => setShowTestMode(true)}
              style={{
                marginTop: 12,
                background: `linear-gradient(135deg, ${C.purple} 0%, #5b21b6 100%)`,
                border: "none",
                color: "white",
                fontSize: 12,
                fontWeight: 700,
                padding: "8px 12px",
                borderRadius: 9,
                cursor: "pointer",
                fontFamily: "inherit",
                width: "100%",
                boxShadow: "0 3px 8px rgba(124,58,237,0.35)",
              }}
            >
              🧪 Mode test
            </button>
          )}
          <button
            onClick={() => setRole(null)}
            style={{
              marginTop: 12,
              background: C.light,
              border: `1px solid ${C.border}`,
              color: "#555",
              fontSize: 12,
              fontWeight: 600,
              padding: "7px 12px",
              borderRadius: 9,
              cursor: "pointer",
              fontFamily: "inherit",
              width: "100%",
            }}
          >
            Déconnexion
          </button>
        </div>
      </div>
      <div className="eco-main" style={{ flex: 1, overflowY: "auto" }}>
        <div className="eco-main-inner" style={{ padding: "26px 30px" }}>
          {previewRole && (
            /* @__PURE__ */ <div
              style={{
                position: "sticky",
                top: 0,
                zIndex: 200,
                marginBottom: 16,
                marginTop: -6,
                borderRadius: 10,
                overflow: "hidden",
                boxShadow: "0 4px 14px rgba(124,58,237,0.35)",
              }}
            >
              <div
                style={{
                  background: `linear-gradient(90deg, ${C.purple} 0%, #5b21b6 100%)`,
                  padding: "10px 16px",
                  display: "flex",
                  alignItems: "center",
                  gap: 10,
                  flexWrap: "wrap",
                }}
              >
                <span style={{ fontSize: 17 }}>🧪</span>
                <span style={{ color: "white", fontWeight: 700, fontSize: 13 }}>
                  Mode test — vous visualisez l'application comme {previewUser?.nom} (
                  {ROLES[previewRole]?.label || previewRole})
                </span>
                <span style={{ color: "rgba(255,255,255,0.85)", fontSize: 11.5 }}>
                  Aucune action ici n'affecte votre propre compte, mais les actions effectuées utilisent vos droits
                  SharePoint réels — soyez prudent(e).
                </span>
                <Btn
                  small={true}
                  variant="light"
                  style={{ marginLeft: "auto" }}
                  onClick={() => {
                    setPreviewRole(null);
                    setPreviewUser(null);
                    setActive(NAV_BY_ROLE[roleAuth][0]);
                  }}
                >
                  ✕ Quitter le mode test
                </Btn>
              </div>
            </div>
          )}
          {role === "etudiant" &&
            (() => {
              const mesPresencesGlobal = (data.presences || []).filter(
                (p) =>
                  (user?.numInscription && p.numInscription === user.numInscription) ||
                  (user?.etudiantId && p.etudiantId === user.etudiantId)
              );
              const parModuleGlobal = {};
              mesPresencesGlobal.forEach((p) => {
                const mod = p.module || "Module non précisé";
                if (!parModuleGlobal[mod]) parModuleGlobal[mod] = 0;
                parModuleGlobal[mod] += SEANCE_DUREE[p.seance || p.heure] || 1.5;
              });
              const alertes = Object.keys(parModuleGlobal)
                .map((m) => {
                  const moduleInfo = (data.cours || []).find((c) => c.intitule === m);
                  const heuresTotal = moduleInfo?.heures || 0;
                  const taux = heuresTotal > 0 ? Math.round((parModuleGlobal[m] / heuresTotal) * 10000) / 100 : 0;
                  return { module: m, taux };
                })
                .filter((a) => a.taux >= 20)
                .sort((a, b) => b.taux - a.taux);
              if (alertes.length === 0) return null;
              const critique = alertes.filter((a) => a.taux >= 25);
              return (
                /* @__PURE__ */ <div
                  style={{
                    marginBottom: 18,
                    borderRadius: 10,
                    overflow: "hidden",
                    border: `1px solid ${critique.length ? "#f3b0a8" : "#f5c98a"}`,
                  }}
                >
                  <div style={{ background: critique.length ? "#fdeeec" : "#fff6e8", padding: "14px 18px" }}>
                    <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 8 }}>
                      <span style={{ fontSize: 20 }}>{critique.length ? "🚨" : "⚠️"}</span>
                      <span style={{ fontWeight: 800, fontSize: 14, color: critique.length ? C.red : "#a5690a" }}>
                        {critique.length ? "Alerte : risque d'ajournement" : "Attention : taux d'absence élevé"}
                      </span>
                    </div>
                    {alertes.map((a) => (
                      /* @__PURE__ */ <div
                        key={a.module}
                        style={{ fontSize: 12.5, color: "#444", padding: "3px 0 3px 30px" }}
                      >
                        <strong>{a.module}</strong> — {a.taux.toFixed(2)}% d'absence du volume horaire
                        {a.taux >= 25 ? (
                          /* @__PURE__ */ <span style={{ color: C.red, fontWeight: 700 }}>
                             — module considéré comme <u>ajourné</u> au-delà de 25%
                          </span>
                        ) : (
                          /* @__PURE__ */ <span style={{ color: "#a5690a" }}>
                             — au-delà de 25%, le module sera considéré comme ajourné
                          </span>
                        )}
                      </div>
                    ))}
                  </div>
                </div>
              );
            })()}
          {["k.bouabid@esrmi.ma", "a.bahalla@esrmi.ma"].includes((user?.email || "").toLowerCase()) &&
            (() => {
              const etudiantsActifs = (data.etudiants || []).filter((e) => e.statut !== "Inactif");
              const sansTarif = etudiantsActifs.filter((e) => !e.tarifGlobalFormation);
              const sansCin = etudiantsActifs.filter((e) => !e.cin || !e.cin.trim());
              if (sansTarif.length === 0 && sansCin.length === 0) return null;
              const morceaux = [];
              if (sansTarif.length > 0) morceaux.push(`${sansTarif.length} sans tarif de formation attribué`);
              if (sansCin.length > 0) morceaux.push(`${sansCin.length} sans CIN renseigné (obligatoire)`);
              return (
                /* @__PURE__ */ <div
                  style={{
                    marginBottom: 18,
                    borderRadius: 10,
                    overflow: "hidden",
                    border: `2px solid ${C.red}`,
                    boxShadow: "0 2px 10px rgba(192,57,43,0.25)",
                  }}
                >
                  <div style={{ background: "#fdeeec", padding: "14px 18px" }}>
                    <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                      <span style={{ fontSize: 22 }}>🚨</span>
                      <span style={{ fontWeight: 800, fontSize: 14, color: C.red }}>
                        Étudiant(s) avec informations manquantes : {morceaux.join(", ")} — merci de compléter leur
                        dossier
                      </span>
                      <Btn
                        small={true}
                        variant="danger"
                        onClick={() => setActive("etudiants")}
                        style={{ marginLeft: "auto" }}
                      >
                        Voir les étudiants
                      </Btn>
                    </div>
                  </div>
                </div>
              );
            })()}
          {["administrateur", "directrice"].includes(role) &&
            (() => {
              const enAttente = (data.demandesDocuments || []).filter(
                (d) => d.statut === "En attente" || d.statut === "En cours"
              );
              if (enAttente.length === 0) return null;
              return (
                /* @__PURE__ */ <div
                  style={{ marginBottom: 18, borderRadius: 10, overflow: "hidden", border: "1px solid #f5c98a" }}
                >
                  <div style={{ background: "#fff6e8", padding: "14px 18px" }}>
                    <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                      <span style={{ fontSize: 20 }}>📨</span>
                      <span style={{ fontWeight: 800, fontSize: 14, color: "#a5690a" }}>
                        {enAttente.length} demande{enAttente.length > 1 ? "s" : ""} de document administratif en attente
                        de traitement
                      </span>
                      <Btn
                        small={true}
                        variant="light"
                        onClick={() => setActive("demandes")}
                        style={{ marginLeft: "auto" }}
                      >
                        Voir les demandes
                      </Btn>
                    </div>
                  </div>
                </div>
              );
            })()}
          {(() => {
            const emailActuel = (user?.email || "").toLowerCase();
            const fiches = data.pointageFiches || [];
            let count = 0,
              message = "",
              statutCible = "";
            if (emailActuel === "i.elmzioui@esrmi.ma") {
              statutCible = "Soumise";
              count = fiches.filter((f) => f.statut === statutCible).length;
              message = `fiche(s) de pointage en attente de validation des séances pédagogiques`;
            } else if (["k.bouabid@esrmi.ma"].includes(emailActuel)) {
              statutCible = "ValideeIliass";
              count = fiches.filter((f) => f.statut === statutCible).length;
              message = `fiche(s) de pointage validée(s) en attente d'accusé de réception et d'impression`;
            } else if (["fz.zniber@esrmi.ma"].includes(emailActuel)) {
              statutCible = "ReceptionAccusee";
              count = fiches.filter((f) => f.statut === statutCible).length;
              message = `fiche(s) de pointage en attente de l'ordre de virement`;
            } else {
              return null;
            }
            if (count === 0) return null;
            return (
              /* @__PURE__ */ <div
                style={{ marginBottom: 18, borderRadius: 10, overflow: "hidden", border: `1px solid ${C.green}66` }}
              >
                <div style={{ background: "#eafaf0", padding: "14px 18px" }}>
                  <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                    <span style={{ fontSize: 20 }}>🕛</span>
                    <span style={{ fontWeight: 800, fontSize: 14, color: C.green }}>
                      {count} {message}
                    </span>
                    <Btn
                      small={true}
                      variant="light"
                      onClick={() => {
                        setActive("pointage");
                        setPointageFiltreCible(statutCible);
                      }}
                      style={{ marginLeft: "auto" }}
                    >
                      Voir le pointage
                    </Btn>
                  </div>
                </div>
              </div>
            );
          })()}
          <Page
            data={data}
            setData={setData}
            role={role}
            user={user}
            toast={toast}
            pointageFiltreCible={pointageFiltreCible}
            setPointageFiltreCible={setPointageFiltreCible}
          />
        </div>
      </div>
      <Toast msg={toastMsg} />
      {showTestMode && (
        /* @__PURE__ */ <Modal title="🧪 Tester en tant que..." onClose={() => setShowTestMode(false)} width={520}>
          <p style={{ fontSize: 12.5, color: C.muted, marginTop: 0, marginBottom: 14 }}>
            Sélectionnez un étudiant ou un professeur pour visualiser l'application exactement comme cette personne la
            verrait, sans avoir besoin de son mot de passe.
          </p>
          <input
            placeholder="🔍  Rechercher par nom ou email..."
            value={testModeSearch}
            onChange={(e) => setTestModeSearch(e.target.value)}
            style={{
              width: "100%",
              padding: "10px 14px",
              border: `1px solid ${C.border}`,
              borderRadius: 8,
              fontSize: 13,
              boxSizing: "border-box",
              fontFamily: "inherit",
              marginBottom: 12,
            }}
          />
          <div style={{ maxHeight: 360, overflowY: "auto", display: "flex", flexDirection: "column", gap: 6 }}>
            {(() => {
              const q = testModeSearch.toLowerCase();
              const candidatsEtudiants = (data.etudiants || [])
                .filter(
                  (e) =>
                    e.statut !== "Inactif" && e.email && `${e.prenom} ${e.nom} ${e.email}`.toLowerCase().includes(q)
                )
                .slice(0, 30);
              const candidatsProfs = (data.professeurs || [])
                .filter((p) => p.email && `${p.prenom} ${p.nom} ${p.email}`.toLowerCase().includes(q))
                .slice(0, 30);
              const candidatsDirection = Object.entries(USER_ROLES)
                .filter(
                  ([email, info]) => info.role === "directrice" && `${info.nom} ${email}`.toLowerCase().includes(q)
                )
                .map(([email, info]) => ({ email, nom: info.nom, titre: info.titre }));
              const choisir = (personne, roleChoisi) => {
                const userObj =
                  roleChoisi === "etudiant"
                    ? {
                        email: personne.email,
                        nom: `${personne.prenom} ${personne.nom}`,
                        titre: "Étudiant",
                        etudiantId: personne.id,
                        numInscription: personne.numInscription,
                        token: userAuth?.token,
                      }
                    : roleChoisi === "professeur"
                      ? {
                          email: personne.email,
                          nom: `${personne.prenom} ${personne.nom}`,
                          titre: "Professeur",
                          professeurId: personne.id,
                          token: userAuth?.token,
                        }
                      : { email: personne.email, nom: personne.nom, titre: personne.titre, token: userAuth?.token };
                setPreviewRole(roleChoisi);
                setPreviewUser(userObj);
                setActive(NAV_BY_ROLE[roleChoisi][0]);
                setShowTestMode(false);
                setTestModeSearch("");
              };
              if (candidatsEtudiants.length === 0 && candidatsProfs.length === 0 && candidatsDirection.length === 0)
                return (
                  /* @__PURE__ */ <div style={{ textAlign: "center", color: C.muted, fontSize: 12.5, padding: 20 }}>
                    Aucun résultat
                  </div>
                );
              return [
                candidatsDirection.length > 0 && (
                  /* @__PURE__ */ <div
                    key="di"
                    style={{
                      fontSize: 10.5,
                      fontWeight: 800,
                      color: C.navy,
                      textTransform: "uppercase",
                      letterSpacing: "0.06em",
                      marginTop: 4,
                    }}
                  >
                    Direction
                  </div>
                ),
                ...candidatsDirection.map((d) => (
                  /* @__PURE__ */ <button
                    key={"d" + d.email}
                    onClick={() => choisir(d, "directrice")}
                    style={{
                      display: "flex",
                      alignItems: "center",
                      gap: 8,
                      textAlign: "left",
                      padding: "8px 10px",
                      border: `1px solid ${C.border}`,
                      borderRadius: 8,
                      background: "white",
                      cursor: "pointer",
                      fontFamily: "inherit",
                    }}
                  >
                    <span style={{ fontSize: 15 }}>🏛️</span>
                    <div>
                      <div style={{ fontSize: 12.5, fontWeight: 700, color: "#232733" }}>{d.nom}</div>
                      <div style={{ fontSize: 10.5, color: C.muted }}>{d.email}</div>
                    </div>
                  </button>
                )),
                candidatsEtudiants.length > 0 && (
                  /* @__PURE__ */ <div
                    key="et"
                    style={{
                      fontSize: 10.5,
                      fontWeight: 800,
                      color: C.orange,
                      textTransform: "uppercase",
                      letterSpacing: "0.06em",
                      marginTop: 8,
                    }}
                  >
                    Étudiants
                  </div>
                ),
                ...candidatsEtudiants.map((e) => (
                  /* @__PURE__ */ <button
                    key={"e" + e.id}
                    onClick={() => choisir(e, "etudiant")}
                    style={{
                      display: "flex",
                      alignItems: "center",
                      gap: 8,
                      textAlign: "left",
                      padding: "8px 10px",
                      border: `1px solid ${C.border}`,
                      borderRadius: 8,
                      background: "white",
                      cursor: "pointer",
                      fontFamily: "inherit",
                    }}
                  >
                    <span style={{ fontSize: 15 }}>🎓</span>
                    <div>
                      <div style={{ fontSize: 12.5, fontWeight: 700, color: "#232733" }}>
                        {e.prenom} {e.nom}
                      </div>
                      <div style={{ fontSize: 10.5, color: C.muted }}>{e.email}</div>
                    </div>
                  </button>
                )),
                candidatsProfs.length > 0 && (
                  /* @__PURE__ */ <div
                    key="pr"
                    style={{
                      fontSize: 10.5,
                      fontWeight: 800,
                      color: C.teal,
                      textTransform: "uppercase",
                      letterSpacing: "0.06em",
                      marginTop: 8,
                    }}
                  >
                    Professeurs
                  </div>
                ),
                ...candidatsProfs.map((p) => (
                  /* @__PURE__ */ <button
                    key={"p" + p.id}
                    onClick={() => choisir(p, "professeur")}
                    style={{
                      display: "flex",
                      alignItems: "center",
                      gap: 8,
                      textAlign: "left",
                      padding: "8px 10px",
                      border: `1px solid ${C.border}`,
                      borderRadius: 8,
                      background: "white",
                      cursor: "pointer",
                      fontFamily: "inherit",
                    }}
                  >
                    <span style={{ fontSize: 15 }}>👨‍🏫</span>
                    <div>
                      <div style={{ fontSize: 12.5, fontWeight: 700, color: "#232733" }}>
                        {p.prenom} {p.nom}
                      </div>
                      <div style={{ fontSize: 10.5, color: C.muted }}>{p.email}</div>
                    </div>
                  </button>
                )),
              ];
            })()}
          </div>
        </Modal>
      )}
      {spLoading && (
        /* @__PURE__ */ <div
          style={{
            position: "fixed",
            bottom: 24,
            left: 24,
            background: C.navy,
            color: "white",
            padding: "10px 18px",
            borderRadius: 11,
            fontSize: 13,
            fontWeight: 500,
            zIndex: 2e3,
            display: "flex",
            alignItems: "center",
            gap: 10,
            boxShadow: C.shadowLg,
          }}
        >
          <div
            style={{
              width: 14,
              height: 14,
              border: "2px solid rgba(255,255,255,0.3)",
              borderTop: "2px solid white",
              borderRadius: "50%",
              animation: "spin 0.8s linear infinite",
            }}
          />
          Chargement des données SharePoint...
        </div>
      )}
      {roleAuth && (!isOnline || pendingSyncCount > 0) && (
        /* @__PURE__ */ <div
          style={{
            position: "fixed",
            bottom: spLoading ? 74 : 24,
            left: 24,
            background: !isOnline ? "#c0392b" : "#e56b2d",
            color: "white",
            padding: "10px 18px",
            borderRadius: 11,
            fontSize: 13,
            fontWeight: 600,
            zIndex: 2e3,
            display: "flex",
            alignItems: "center",
            gap: 10,
            boxShadow: C.shadowLg,
            maxWidth: 340,
          }}
        >
          <span style={{ fontSize: 16 }}>{!isOnline ? "🔌" : syncingQueue ? "🔄" : "⏳"}</span>
          <span>
            {!isOnline
              ? `Hors ligne — vos modifications sont enregistrées et seront synchronisées au retour de la connexion${pendingSyncCount > 0 ? ` (${pendingSyncCount} en attente)` : ""}`
              : syncingQueue
                ? "Synchronisation en cours..."
                : `${pendingSyncCount} modification(s) en attente de synchronisation`}
          </span>
          {isOnline && !syncingQueue && pendingSyncCount > 0 && userAuth?.token && (
            <button
              onClick={() => runSync(userAuth.token)}
              style={{
                border: "none",
                background: "rgba(255,255,255,0.22)",
                color: "white",
                borderRadius: 8,
                padding: "5px 10px",
                fontSize: 12,
                fontWeight: 700,
                cursor: "pointer",
                fontFamily: "inherit",
                flexShrink: 0,
              }}
            >
              🔄 Synchroniser
            </button>
          )}
        </div>
      )}
      <style>{`@keyframes spin{to{transform:rotate(360deg)}}`}</style>
    </div>
  );
}
