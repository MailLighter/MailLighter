/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    
    // Associer chaque bouton à sa fonction
    document.getElementById("remove-images").onclick = removeImages;
    document.getElementById("remove-attachments").onclick = removeAttachments;
    document.getElementById("keep-two-replies").onclick = keepTwoReplies;
    document.getElementById("clean-all").onclick = showCleanAllConfirmation;
  }
});

// Fonction pour afficher un message de statut
function showMessage(message, isSuccess = true) {
  const statusElement = document.getElementById("status-message");
  statusElement.innerText = message;
  statusElement.style.color = isSuccess ? "green" : "red";
}

// Affiche une confirmation personnalisée avant d'exécuter cleanAll
function showCleanAllConfirmation() {
  const statusElement = document.getElementById("status-message");
  statusElement.innerHTML =
    '<span style="color: orange;">⚠️ Êtes-vous sûr de vouloir tout supprimer ?</span><br>' +
    '<button id="confirm-clean-all-yes">Oui</button> ' +
    '<button id="confirm-clean-all-no">Non</button>';
  document.getElementById("confirm-clean-all-yes").onclick = async function() {
    statusElement.innerText = "Nettoyage en cours...";
    await cleanAll();
  };
  document.getElementById("confirm-clean-all-no").onclick = function() {
    statusElement.innerText = "Action annulée.";
    statusElement.style.color = "gray";
  };
}

// ========================================
// FONCTION 1 : Supprimer les images
// ========================================
async function removeImages() {
  try {
    const item = Office.context.mailbox.item;
    
    // Vérifier si on peut modifier le corps
    if (!item.body || !item.body.setAsync) {
      showMessage("❌ Modification du corps non disponible en mode lecture", false);
      return;
    }
    
    // Accédez au corps du mail
    item.body.getAsync(Office.CoercionType.Html, function(result) {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        let htmlBody = result.value;
        
        // Extraire toutes les images et calculer la taille totale estimée
        const imgMatches = htmlBody.match(/<img[^>]*>/gi) || [];
        const totalSize = calculateImageSize(imgMatches);
        
        // Supprimer tous les tags <img>
        htmlBody = htmlBody.replace(/<img[^>]*>/gi, "");
        
        // Remplacer le corps du mail
        item.body.setAsync(htmlBody, { coercionType: Office.CoercionType.Html }, function(setResult) {
          if (setResult.status === Office.AsyncResultStatus.Succeeded) {
            const sizeText = totalSize > 0 ? " (" + formatFileSize(totalSize) + ")" : "";
            showMessage("✅ " + imgMatches.length + " image(s) supprimée(s) !" + sizeText, true);
          } else {
            showMessage("❌ Erreur lors de la suppression des images: " + setResult.error.message, false);
          }
        });
      } else {
        showMessage("❌ Erreur de lecture du mail", false);
      }
    });
  } catch (error) {
    showMessage("❌ Erreur: " + error.message, false);
  }
}

// Fonction utilitaire : calcule la taille estimée des images
function calculateImageSize(imgMatches) {
  let totalSize = 0;
  imgMatches.forEach(function(imgTag) {
    // Chercher l'attribut data-size si présent
    const dataSizeMatch = imgTag.match(/data-size="?(\d+)"?/i);
    if (dataSizeMatch) {
      totalSize += parseInt(dataSizeMatch[1], 10);
    } else {
      // Sinon, estimer basé sur width et height
      const widthMatch = imgTag.match(/width="?(\d+)"?/i);
      const heightMatch = imgTag.match(/height="?(\d+)"?/i);
      if (widthMatch && heightMatch) {
        const width = parseInt(widthMatch[1], 10);
        const height = parseInt(heightMatch[1], 10);
        // Estimation : ~5KB par 100x100 pixels
        const estimatedSize = Math.round((width * height) / 2000) * 5120;
        totalSize += estimatedSize;
      } else {
        // Default : 50KB par image si pas de dimensions
        totalSize += 51200;
      }
    }
  });
  return totalSize;
}

// Fonction utilitaire : formate les octets en KB/MB (minimum KB)
function formatFileSize(bytes) {
  if (bytes === 0) return "0 KB";
  const k = 1024;
  let sizeInKB = bytes / k;
  
  if (sizeInKB < 1) {
    return "< 1 KB économisé(s)";
  } else if (sizeInKB < 1024) {
    return Math.round(sizeInKB * 100) / 100 + " KB économisé(s)";
  } else if (sizeInKB < 1024 * 1024) {
    return Math.round((sizeInKB / 1024) * 100) / 100 + " MB économisé(s)";
  } else {
    return Math.round((sizeInKB / (1024 * 1024)) * 100) / 100 + " GB économisé(s)";
  }
}

// ========================================
// FONCTION 2 : Supprimer les pièces jointes
// ========================================
async function removeAttachments() {
  try {
    const item = Office.context.mailbox.item;
    const callId = Date.now() + Math.random(); // Identifiant unique pour cet appel
    
    // En mode composition, utiliser getAttachmentsAsync
    if (item.itemType === Office.MailboxEnums.ItemType.Message && typeof item.getAttachmentsAsync === 'function') {
      item.getAttachmentsAsync(function(result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          const attachments = result.value;
          
          if (!attachments || attachments.length === 0) {
            showMessage("ℹ️ Aucune pièce jointe à supprimer", true);
            return;
          }
          
          const attachmentCount = attachments.length;
          
          // Calculer la taille totale des pièces jointes
          let totalSize = 0;
          attachments.forEach(function(att) {
            if (att.size) {
              totalSize += att.size;
            }
          });
          
          // Supprimer chaque pièce jointe
          let successCount = 0;
          attachments.forEach(function(attachment) {
            item.removeAttachmentAsync(attachment.id, function(removeResult) {
              if (removeResult.status === Office.AsyncResultStatus.Succeeded) {
                successCount++;
                // Vérifier qu'on a bien le bon nombre et que c'est le bon appel
                if (successCount === attachmentCount && attachmentCount > 0) {
                  const sizeText = totalSize > 0 ? " (" + formatFileSize(totalSize) + ")" : "";
                  showMessage("✅ " + successCount + " pièce(s) jointe(s) supprimée(s)!" + sizeText, true);
                }
              } else {
                showMessage("❌ Erreur suppression PJ: " + removeResult.error.message, false);
              }
            });
          });
        } else {
          showMessage("❌ Erreur récupération PJ: " + result.error.message, false);
        }
      });
    } else {
      // Mode lecture
      if (!item.attachments || !Array.isArray(item.attachments) || item.attachments.length === 0) {
        showMessage("ℹ️ Aucune pièce jointe à supprimer", true);
        return;
      }
      
      const attachmentCount = item.attachments.length;
      
      // Calculer la taille totale des pièces jointes
      let totalSize = 0;
      item.attachments.forEach(function(att) {
        if (att.size) {
          totalSize += att.size;
        }
      });
      
      const attachmentIds = item.attachments.map(att => att.id);
      let successCount = 0;
      attachmentIds.forEach(function(attachmentId) {
        item.attachments.removeAttachmentAsync(attachmentId, function(result) {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            successCount++;
            // Vérifier qu'on a bien le bon nombre et que c'est le bon appel
            if (successCount === attachmentCount && attachmentCount > 0) {
              const sizeText = totalSize > 0 ? " (" + formatFileSize(totalSize) + ")" : "";
              showMessage("✅ " + successCount + " pièce(s) jointe(s) supprimée(s)!" + sizeText, true);
            }
          }
        });
      });
    }
  } catch (error) {
    showMessage("❌ Erreur: " + error.message, false);
  }
}

// ========================================
// FONCTION 3 : Garder les 2 derniers replies
// ========================================
async function keepTwoReplies() {
  try {
    const item = Office.context.mailbox.item;
    item.body.getAsync(Office.CoercionType.Text, function(result) {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        showMessage("❌ Erreur de lecture du mail", false);
        return;
      }
      const textBody = result.value || "";

      // Chercher tous les "De :" (qui marquent le début d'un reply)
      const deMatches = [];
      const deRegex = /De\s*:/gi;
      let match;
      while ((match = deRegex.exec(textBody)) !== null) {
        deMatches.push(match.index);
      }

      if (deMatches.length === 0) {
        showMessage("ℹ️ Aucun reply détecté dans le mail", true);
        return;
      }

      if (deMatches.length <= 2) {
        showMessage("ℹ️ " + deMatches.length + " reply(s) trouvé(s), moins de 3 donc aucun changement", true);
        return;
      }

      // Construire les segments : du 1er "De :" au 2e "De :" au 3e "De :", etc.
      const segments = [];
      for (let i = 0; i < deMatches.length; i++) {
        const start = deMatches[i];
        const end = (i + 1 < deMatches.length) ? deMatches[i + 1] : textBody.length;
        segments.push(textBody.substring(start, end));
      }

      // Le corps jusqu'au premier "De :"
      const firstReplyPos = deMatches[0];
      let cleanedText = textBody.substring(0, firstReplyPos);

      // Ajouter les 2 premiers segments (2 replies)
      cleanedText += (segments[0] || "");
      if (segments.length > 1) {
        cleanedText += (segments[1] || "");
      }

      // Remplacer le corps
      item.body.setAsync(cleanedText, { coercionType: Office.CoercionType.Text }, function(setResult) {
        if (setResult.status === Office.AsyncResultStatus.Succeeded) {
          showMessage("✅ Replies nettoyés (" + deMatches.length + " trouvés, 2 conservés)!", true);
        } else {
          showMessage("❌ Erreur lors du nettoyage des replies: " + (setResult.error && setResult.error.message ? setResult.error.message : 'unknown'), false);
        }
      });
    });
  } catch (error) {
    showMessage("❌ Erreur: " + error.message, false);
  }
}

// ========================================
// FONCTION 4 : Tout supprimer (réutilise les fonctions existantes)
// ========================================
async function cleanAll() {
  try {
    showMessage("Nettoyage en cours...", true);
    await removeImages();
    await removeAttachments();
    await keepTwoReplies();
    showMessage("✅ Mail complètement nettoyé!", true);
  } catch (error) {
    showMessage("❌ Erreur: " + error.message, false);
  }
}
