/* global Office, Word */

const STORAGE_KEY_API = "gemini_word_api_key";
const STORAGE_KEY_MODEL = "gemini_word_model";
const STORAGE_KEY_SYS = "gemini_word_system_prompt";

// 1. PANNEAU D'OUTILS CHIRURGICAUX WORD EXPOSÉS À GEMINI
const WORD_TOOLS = [{
  functionDeclarations: [
    {
      name: "lire_selection",
      description: "Lit le texte actuellement sélectionné par l'utilisateur dans le document Word."
    },
    {
      name: "lire_document_complet",
      description: "Lit l'intégralité du texte du document Word actif."
    },
    {
      name: "lire_plan_document",
      description: "Extrait la table des matières structurelle (tous les paragraphes ayant un style Titre 1, Titre 2, Titre 3) avec leur index pour comprendre l'architecture d'un long document."
    },
    {
      name: "inserer_contenu_riche",
      description: "Insère ou remplace du contenu dans Word avec de vrais styles natifs Word ET du formatage intra-paragraphe (**gras**, *italique*, ==surligné==).",
      parameters: {
        type: "OBJECT",
        properties: {
          mode: {
            type: "STRING",
            enum: ["replace", "after", "before", "end"],
            description: "'replace' remplace la sélection, 'after' insère après le curseur/sélection, 'before' insère avant, 'end' ajoute à la fin du document."
          },
          blocs: {
            type: "ARRAY",
            description: "Liste ordonnée des paragraphes à insérer.",
            items: {
              type: "OBJECT",
              properties: {
                style: {
                  type: "STRING",
                  enum: ["Title", "Subtitle", "Heading1", "Heading2", "Heading3", "Normal", "ListBullet", "Quote", "IntenseQuote"],
                  description: "Style natif Word à appliquer au paragraphe."
                },
                texte: {
                  type: "STRING",
                  description: "Contenu du paragraphe. Utilise **mot** pour le gras, *mot* pour l'italique, et ==mot== pour le surlignage jaune."
                },
                alignement: {
                  type: "STRING",
                  enum: ["Left", "Centered", "Right", "Justified"],
                  description: "Alignement optionnel du paragraphe (par défaut Left, ou Justified pour un rapport)."
                }
              },
              required: ["style", "texte"]
            }
          }
        },
        required: ["mode", "blocs"]
      }
    },
    {
      name: "rechercher_remplacer",
      description: "Recherche une expression exacte dans tout le document (ou dans la sélection) et la remplace par un nouveau texte, avec mise en forme optionnelle.",
      parameters: {
        type: "OBJECT",
        properties: {
          texte_recherche: { type: "STRING", description: "Le texte exact à trouver dans Word." },
          texte_remplacement: { type: "STRING", description: "Le nouveau texte de remplacement." },
          casse_exacte: { type: "BOOLEAN", description: "Respecter les majuscules/minuscules." },
          mettre_en_gras: { type: "BOOLEAN", description: "Appliquer le gras au texte remplacé." },
          surligner_jaune: { type: "BOOLEAN", description: "Surligner en jaune le texte remplacé." }
        },
        required: ["texte_recherche", "texte_remplacement"]
      }
    },
    {
      name: "ajouter_commentaire",
      description: "Ajoute un vrai commentaire Word dans la marge sur la sélection actuelle ou sur un passage précis du document (idéal pour la relecture critique sans écraser le texte).",
      parameters: {
        type: "OBJECT",
        properties: {
          passage_cible: {
            type: "STRING",
            description: "Texte exact dans le document auquel attacher le commentaire. Laisser vide pour attacher à la sélection actuelle."
          },
          commentaire: {
            type: "STRING",
            description: "Contenu de la note ou suggestion dans la marge."
          }
        },
        required: ["commentaire"]
      }
    },
    {
      name: "inserer_tableau",
      description: "Insère un vrai tableau Word natif structuré après la sélection actuelle.",
      parameters: {
        type: "OBJECT",
        properties: {
          lignes: {
            type: "ARRAY",
            description: "Tableau 2D (la première ligne sert d'en-tête).",
            items: {
              type: "ARRAY",
              items: { type: "STRING" }
            }
          }
        },
        required: ["lignes"]
      }
    },
    {
      name: "gerer_mise_en_page",
      description: "Insère un saut de page, un saut de section, ou configure le texte de l'en-tête et du pied de page du document.",
      parameters: {
        type: "OBJECT",
        properties: {
          action: {
            type: "STRING",
            enum: ["saut_de_page", "definir_entete", "definir_pied_de_page"],
            description: "Action de mise en page à exécuter."
          },
          texte: {
            type: "STRING",
            description: "Texte à placer dans l'en-tête ou le pied de page (requis si action est definir_entete ou definir_pied_de_page)."
          }
        },
        required: ["action"]
      }
    }
  ]
}];

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    initUI();
  }
});

function initUI() {
  const apiKeyInput = document.getElementById("apiKeyInput");
  const modelSelect = document.getElementById("modelSelect");
  const sysInput = document.getElementById("systemInstructionInput");
  const settingsPanel = document.getElementById("settingsPanel");

  const savedKey = localStorage.getItem(STORAGE_KEY_API) || "";
  const savedModel = localStorage.getItem(STORAGE_KEY_MODEL) || "gemini-4-hydrogen";
  const savedSys = localStorage.getItem(STORAGE_KEY_SYS) || "";

  apiKeyInput.value = savedKey;
  modelSelect.value = savedModel;
  sysInput.value = savedSys;

  if (!savedKey) {
    settingsPanel.classList.remove("hidden");
  }

  document.getElementById("toggleSettingsBtn").addEventListener("click", () => {
    settingsPanel.classList.toggle("hidden");
  });

  document.getElementById("saveKeyBtn").addEventListener("click", () => {
    localStorage.setItem(STORAGE_KEY_API, apiKeyInput.value.trim());
    localStorage.setItem(STORAGE_KEY_SYS, sysInput.value.trim());
    settingsPanel.classList.add("hidden");
    setStatus("Paramètres enregistrés.");
  });

  modelSelect.addEventListener("change", () => {
    localStorage.setItem(STORAGE_KEY_MODEL, modelSelect.value);
    setStatus(`Modèle actif : ${modelSelect.options[modelSelect.selectedIndex].text}`);
  });

  document.querySelectorAll(".chip").forEach((chip) => {
    chip.addEventListener("click", () => {
      const prompt = chip.getAttribute("data-prompt");
      document.getElementById("userPrompt").value = prompt;
      runAgent();
    });
  });

  document.getElementById("sendBtn").addEventListener("click", runAgent);
  document.getElementById("userPrompt").addEventListener("keydown", (e) => {
    if (e.key === "Enter" && !e.shiftKey) {
      e.preventDefault();
      runAgent();
    }
  });
}

function setStatus(text) {
  document.getElementById("statusBar").innerText = text;
}

function appendMessage(text, type = "agent") {
  const history = document.getElementById("chatHistory");
  const div = document.createElement("div");
  div.className = `msg msg-${type}`;
  div.innerText = text;
  history.appendChild(div);
  history.scrollTop = history.scrollHeight;
}

// Helper : Applique le formatage riche (**gras**, *italique*, ==surligné==) dans un paragraphe Word
async function applyInlineFormatting(context, paragraph, rawText) {
  // Nettoyer d'abord les marqueurs dans le paragraphe
  const cleanText = rawText
    .replace(/\*\*(.+?)\*\*/g, "$1")
    .replace(/\*(.+?)\*/g, "$1")
    .replace(/==(.+?)==/g, "$1");

  paragraph.insertText(cleanText, Word.InsertLocation.replace);
  await context.sync();

  // Appliquer le gras sur les segments **...**
  const boldMatches = [...rawText.matchAll(/\*\*(.+?)\*\*/g)];
  for (const m of boldMatches) {
    const results = paragraph.search(m[1], { matchCase: true });
    results.load("items");
    await context.sync();
    results.items.forEach((r) => { r.font.bold = true; });
  }

  // Appliquer l'italique sur les segments *...*
  const italicMatches = [...rawText.matchAll(/(?<!\*)\*(?!\*)(.+?)(?<!\*)\*(?!\*)/g)];
  for (const m of italicMatches) {
    const results = paragraph.search(m[1], { matchCase: true });
    results.load("items");
    await context.sync();
    results.items.forEach((r) => { r.font.italic = true; });
  }

  // Appliquer le surlignage jaune sur les segments ==...==
  const highlightMatches = [...rawText.matchAll(/==(.+?)==/g)];
  for (const m of highlightMatches) {
    const results = paragraph.search(m[1], { matchCase: true });
    results.load("items");
    await context.sync();
    results.items.forEach((r) => { r.font.highlightColor = "Yellow"; });
  }
}

// 2. EXÉCUTEUR DES OUTILS WORD (OFFICE.JS)
async function executeWordTool(name, args) {
  const useTrackChanges = document.getElementById("trackChangesToggle").checked;

  return await Word.run(async (context) => {
    if (useTrackChanges) {
      context.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
    }

    if (name === "lire_selection") {
      const selection = context.document.getSelection();
      selection.load("text");
      await context.sync();
      return { texte_selectionne: selection.text || "(Aucune sélection active)" };
    }

    if (name === "lire_document_complet") {
      const body = context.document.body;
      body.load("text");
      await context.sync();
      return { texte_document: body.text || "(Document vide)" };
    }

    if (name === "lire_plan_document") {
      const paragraphs = context.document.body.paragraphs;
      paragraphs.load("items/text,items/styleBuiltIn");
      await context.sync();
      const plan = [];
      paragraphs.items.forEach((p, idx) => {
        if (["Title", "Subtitle", "Heading1", "Heading2", "Heading3"].includes(p.styleBuiltIn)) {
          plan.push({ index: idx, niveau: p.styleBuiltIn, titre: p.text.trim() });
        }
      });
      return { plan_document: plan.length ? plan : "Aucun titre structuré détecté." };
    }

    if (name === "inserer_contenu_riche") {
      const { mode, blocs } = args;
      let targetRange;
      let insertLocation = Word.InsertLocation.after;

      if (mode === "end") {
        targetRange = context.document.body.getRange(Word.RangeLocation.end);
      } else {
        targetRange = context.document.getSelection();
        if (mode === "replace") {
          targetRange.clear();
        } else if (mode === "before") {
          insertLocation = Word.InsertLocation.before;
        }
      }

      const styleMap = {
        Title: Word.BuiltInStyleName.title,
        Subtitle: Word.BuiltInStyleName.subtitle,
        Heading1: Word.BuiltInStyleName.heading1,
        Heading2: Word.BuiltInStyleName.heading2,
        Heading3: Word.BuiltInStyleName.heading3,
        ListBullet: Word.BuiltInStyleName.listBullet,
        Quote: Word.BuiltInStyleName.quote,
        IntenseQuote: Word.BuiltInStyleName.intenseQuote,
        Normal: Word.BuiltInStyleName.normal
      };

      for (const bloc of blocs) {
        const para = targetRange.insertParagraph("", insertLocation);
        para.styleBuiltIn = styleMap[bloc.style] || Word.BuiltInStyleName.normal;

        if (bloc.alignement && Word.Alignment[bloc.alignement.toLowerCase()]) {
          para.alignment = Word.Alignment[bloc.alignement.toLowerCase()];
        }

        await applyInlineFormatting(context, para, bloc.texte);
        targetRange = para.getRange();
        insertLocation = Word.InsertLocation.after;
      }

      await context.sync();
      return { statut: "Succès", paragraphes_inseres: blocs.length };
    }

    if (name === "rechercher_remplacer") {
      const { texte_recherche, texte_remplacement, casse_exacte, mettre_en_gras, surligner_jaune } = args;
      const results = context.document.body.search(texte_recherche, { matchCase: !!casse_exacte });
      results.load("items");
      await context.sync();

      for (const item of results.items) {
        item.insertText(texte_remplacement, Word.InsertLocation.replace);
        if (mettre_en_gras) item.font.bold = true;
        if (surligner_jaune) item.font.highlightColor = "Yellow";
      }
      await context.sync();
      return { statut: "Remplacement terminé", occurrences_modifiees: results.items.length };
    }

    if (name === "ajouter_commentaire") {
      const { passage_cible, commentaire } = args;
      let rangeToComment;

      if (passage_cible && passage_cible.trim() !== "") {
        const searchResults = context.document.body.search(passage_cible.trim(), { matchCase: false });
        searchResults.load("items");
        await context.sync();
        if (searchResults.items.length === 0) {
          return { erreur: `Passage introuvable pour le commentaire : "${passage_cible}"` };
        }
        rangeToComment = searchResults.items[0];
      } else {
        rangeToComment = context.document.getSelection();
      }

      rangeToComment.insertComment(commentaire);
      await context.sync();
      return { statut: "Commentaire ajouté dans la marge." };
    }

    if (name === "inserer_tableau") {
      const { lignes } = args;
      if (!lignes || lignes.length === 0) return { erreur: "Tableau vide" };

      const rowCount = lignes.length;
      const colCount = lignes[0].length;
      const selection = context.document.getSelection();
      const table = selection.insertTable(rowCount, colCount, Word.InsertLocation.after, lignes);
      table.styleBuiltIn = Word.BuiltInStyleName.gridTable4Accent1;
      await context.sync();
      return { statut: "Tableau inséré", lignes: rowCount, colonnes: colCount };
    }

    if (name === "gerer_mise_en_page") {
      const { action, texte } = args;
      if (action === "saut_de_page") {
        context.document.body.insertBreak(Word.BreakType.page, Word.InsertLocation.end);
        await context.sync();
        return { statut: "Saut de page inséré." };
      }
      if (action === "definir_entete") {
        const header = context.document.sections.getFirst().getHeader(Word.HeaderFooterType.primary);
        header.clear();
        header.insertParagraph(texte || "", Word.InsertLocation.start);
        await context.sync();
        return { statut: "En-tête mis à jour." };
      }
      if (action === "definir_pied_de_page") {
        const footer = context.document.sections.getFirst().getFooter(Word.HeaderFooterType.primary);
        footer.clear();
        footer.insertParagraph(texte || "", Word.InsertLocation.start);
        await context.sync();
        return { statut: "Pied de page mis à jour." };
      }
    }

    return { erreur: `Outil inconnu : ${name}` };
  });
}

// 3. BOUCLE AUTONOME DE L'AGENT ARCHITECTE WORD
async function runAgent() {
  const apiKey = document.getElementById("apiKeyInput").value.trim();
  const model = document.getElementById("modelSelect").value;
  const userPromptEl = document.getElementById("userPrompt");
  const promptText = userPromptEl.value.trim();
  const sendBtn = document.getElementById("sendBtn");

  if (!apiKey) {
    document.getElementById("settingsPanel").classList.remove("hidden");
    setStatus("Erreur : Clé API manquante.");
    return;
  }
  if (!promptText) return;

  userPromptEl.value = "";
  appendMessage(promptText, "user");
  sendBtn.disabled = true;
  setStatus(`Réflexion (${model})...`);

  const customSys = document.getElementById("systemInstructionInput").value.trim();
  const baseSystemPrompt = `Tu es un Maître Architecte Documentaire et agent IA autonome intégré dans Microsoft Word.
Tu maîtrises parfaitement toutes les fonctionnalités de Word grâce à tes outils :
- Lecture : 'lire_selection', 'lire_document_complet', 'lire_plan_document'.
- Écriture & Mise en forme : 'inserer_contenu_riche' (gère Title, Subtitle, Heading1..3, Normal, ListBullet, Quote, IntenseQuote + formatage intra-texte **gras**, *italique*, ==surligné== + alignement).
- Édition chirurgicale : 'rechercher_remplacer', 'ajouter_commentaire' (dans la marge), 'inserer_tableau', 'gerer_mise_en_page' (sauts de page, en-têtes, pieds de page).

Règles d'or :
1. LIRE AVANT D'AGIR : Avant de modifier, commenter ou restructurer, appelle toujours 'lire_selection', 'lire_plan_document' ou 'lire_document_complet' pour connaître le contexte exact.
2. TYPOLOGIE & HIÉRARCHIE : Respecte la typographie française (espaces insécables avant : ; ! ?, guillemets « ») et ne saute jamais un niveau de titre (Heading1 -> Heading2 -> Heading3).
3. FORMATAGE RICHE : Mets systématiquement en valeur les concepts clés avec **gras**, les termes étrangers ou citations avec *italique*, et utilise ==surligné== uniquement pour les points d'attention ou alertes demandés.
4. NON-DESTRUCTION : Si l'utilisateur demande une critique, une relecture ou des suggestions sur un texte existant sans demander de le remplacer, privilégie 'ajouter_commentaire' dans la marge.
5. Après l'exécution des outils, confirme brièvement tes actions dans le chat.
${customSys ? "\nInstructions spécifiques de l'utilisateur : " + customSys : ""}`;

  const contents = [
    { role: "user", parts: [{ text: promptText }] }
  ];

  try {
    let finished = false;
    let stepCount = 0;

    while (!finished && stepCount < 8) {
      stepCount++;
      const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`;
      const response = await fetch(url, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          systemInstruction: { parts: [{ text: baseSystemPrompt }] },
          contents: contents,
          tools: WORD_TOOLS
        })
      });

      const data = await response.json();
      if (data.error) {
        throw new Error(data.error.message || "Erreur API Gemini");
      }

      const candidate = data.candidates?.[0];
      const modelParts = candidate?.content?.parts || [];
      contents.push({ role: "model", parts: modelParts });

      const functionCalls = modelParts.filter((p) => p.functionCall);

      if (functionCalls.length > 0) {
        const functionResponses = [];
        for (const part of functionCalls) {
          const { name, args } = part.functionCall;
          appendMessage(`⚙️ ${name}...`, "tool");
          setStatus(`Action Word : ${name}...`);

          const toolResult = await executeWordTool(name, args || {});
          functionResponses.push({
            functionResponse: {
              name: name,
              response: toolResult
            }
          });
        }
        contents.push({ role: "user", parts: functionResponses });
      } else {
        const finalText = modelParts.map((p) => p.text || "").join("").trim();
        if (finalText) {
          appendMessage(finalText, "agent");
        }
        finished = true;
      }
    }

    setStatus("Terminé");
  } catch (err) {
    appendMessage(`❌ Erreur : ${err.message}`, "agent");
    setStatus("Échec de l'exécution");
  } finally {
    sendBtn.disabled = false;
  }
}