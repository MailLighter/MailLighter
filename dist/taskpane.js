/******/ (function() { // webpackBootstrap
/******/ 	var __webpack_modules__ = ({

/***/ "./src/taskpane/taskpane.css":
/*!***********************************!*\
  !*** ./src/taskpane/taskpane.css ***!
  \***********************************/
/***/ (function(module, __unused_webpack_exports, __webpack_require__) {

"use strict";
module.exports = __webpack_require__.p + "1fda685b81e1123773f6.css";

/***/ })

/******/ 	});
/************************************************************************/
/******/ 	// The module cache
/******/ 	var __webpack_module_cache__ = {};
/******/ 	
/******/ 	// The require function
/******/ 	function __webpack_require__(moduleId) {
/******/ 		// Check if module is in cache
/******/ 		var cachedModule = __webpack_module_cache__[moduleId];
/******/ 		if (cachedModule !== undefined) {
/******/ 			return cachedModule.exports;
/******/ 		}
/******/ 		// Check if module exists (development only)
/******/ 		if (__webpack_modules__[moduleId] === undefined) {
/******/ 			var e = new Error("Cannot find module '" + moduleId + "'");
/******/ 			e.code = 'MODULE_NOT_FOUND';
/******/ 			throw e;
/******/ 		}
/******/ 		// Create a new module (and put it into the cache)
/******/ 		var module = __webpack_module_cache__[moduleId] = {
/******/ 			// no module.id needed
/******/ 			// no module.loaded needed
/******/ 			exports: {}
/******/ 		};
/******/ 	
/******/ 		// Execute the module function
/******/ 		__webpack_modules__[moduleId](module, module.exports, __webpack_require__);
/******/ 	
/******/ 		// Return the exports of the module
/******/ 		return module.exports;
/******/ 	}
/******/ 	
/******/ 	// expose the modules object (__webpack_modules__)
/******/ 	__webpack_require__.m = __webpack_modules__;
/******/ 	
/************************************************************************/
/******/ 	/* webpack/runtime/global */
/******/ 	!function() {
/******/ 		__webpack_require__.g = (function() {
/******/ 			if (typeof globalThis === 'object') return globalThis;
/******/ 			try {
/******/ 				return this || new Function('return this')();
/******/ 			} catch (e) {
/******/ 				if (typeof window === 'object') return window;
/******/ 			}
/******/ 		})();
/******/ 	}();
/******/ 	
/******/ 	/* webpack/runtime/hasOwnProperty shorthand */
/******/ 	!function() {
/******/ 		__webpack_require__.o = function(obj, prop) { return Object.prototype.hasOwnProperty.call(obj, prop); }
/******/ 	}();
/******/ 	
/******/ 	/* webpack/runtime/make namespace object */
/******/ 	!function() {
/******/ 		// define __esModule on exports
/******/ 		__webpack_require__.r = function(exports) {
/******/ 			if(typeof Symbol !== 'undefined' && Symbol.toStringTag) {
/******/ 				Object.defineProperty(exports, Symbol.toStringTag, { value: 'Module' });
/******/ 			}
/******/ 			Object.defineProperty(exports, '__esModule', { value: true });
/******/ 		};
/******/ 	}();
/******/ 	
/******/ 	/* webpack/runtime/publicPath */
/******/ 	!function() {
/******/ 		var scriptUrl;
/******/ 		if (__webpack_require__.g.importScripts) scriptUrl = __webpack_require__.g.location + "";
/******/ 		var document = __webpack_require__.g.document;
/******/ 		if (!scriptUrl && document) {
/******/ 			if (document.currentScript && document.currentScript.tagName.toUpperCase() === 'SCRIPT')
/******/ 				scriptUrl = document.currentScript.src;
/******/ 			if (!scriptUrl) {
/******/ 				var scripts = document.getElementsByTagName("script");
/******/ 				if(scripts.length) {
/******/ 					var i = scripts.length - 1;
/******/ 					while (i > -1 && (!scriptUrl || !/^http(s?):/.test(scriptUrl))) scriptUrl = scripts[i--].src;
/******/ 				}
/******/ 			}
/******/ 		}
/******/ 		// When supporting browsers where an automatic publicPath is not supported you must specify an output.publicPath manually via configuration
/******/ 		// or pass an empty string ("") and set the __webpack_public_path__ variable from your code to use your own logic.
/******/ 		if (!scriptUrl) throw new Error("Automatic publicPath is not supported in this browser");
/******/ 		scriptUrl = scriptUrl.replace(/^blob:/, "").replace(/#.*$/, "").replace(/\?.*$/, "").replace(/\/[^\/]+$/, "/");
/******/ 		__webpack_require__.p = scriptUrl;
/******/ 	}();
/******/ 	
/******/ 	/* webpack/runtime/jsonp chunk loading */
/******/ 	!function() {
/******/ 		__webpack_require__.b = (typeof document !== 'undefined' && document.baseURI) || self.location.href;
/******/ 		
/******/ 		// object to store loaded and loading chunks
/******/ 		// undefined = chunk not loaded, null = chunk preloaded/prefetched
/******/ 		// [resolve, reject, Promise] = chunk loading, 0 = chunk loaded
/******/ 		var installedChunks = {
/******/ 			"taskpane": 0
/******/ 		};
/******/ 		
/******/ 		// no chunk on demand loading
/******/ 		
/******/ 		// no prefetching
/******/ 		
/******/ 		// no preloaded
/******/ 		
/******/ 		// no HMR
/******/ 		
/******/ 		// no HMR manifest
/******/ 		
/******/ 		// no on chunks loaded
/******/ 		
/******/ 		// no jsonp function
/******/ 	}();
/******/ 	
/************************************************************************/
var __webpack_exports__ = {};
// This entry needs to be wrapped in an IIFE because it needs to be isolated against other entry modules.
!function() {
/*!**********************************!*\
  !*** ./src/taskpane/taskpane.js ***!
  \**********************************/
function _regenerator() { /*! regenerator-runtime -- Copyright (c) 2014-present, Facebook, Inc. -- license (MIT): https://github.com/babel/babel/blob/main/packages/babel-helpers/LICENSE */ var e, t, r = "function" == typeof Symbol ? Symbol : {}, n = r.iterator || "@@iterator", o = r.toStringTag || "@@toStringTag"; function i(r, n, o, i) { var c = n && n.prototype instanceof Generator ? n : Generator, u = Object.create(c.prototype); return _regeneratorDefine2(u, "_invoke", function (r, n, o) { var i, c, u, f = 0, p = o || [], y = !1, G = { p: 0, n: 0, v: e, a: d, f: d.bind(e, 4), d: function d(t, r) { return i = t, c = 0, u = e, G.n = r, a; } }; function d(r, n) { for (c = r, u = n, t = 0; !y && f && !o && t < p.length; t++) { var o, i = p[t], d = G.p, l = i[2]; r > 3 ? (o = l === n) && (u = i[(c = i[4]) ? 5 : (c = 3, 3)], i[4] = i[5] = e) : i[0] <= d && ((o = r < 2 && d < i[1]) ? (c = 0, G.v = n, G.n = i[1]) : d < l && (o = r < 3 || i[0] > n || n > l) && (i[4] = r, i[5] = n, G.n = l, c = 0)); } if (o || r > 1) return a; throw y = !0, n; } return function (o, p, l) { if (f > 1) throw TypeError("Generator is already running"); for (y && 1 === p && d(p, l), c = p, u = l; (t = c < 2 ? e : u) || !y;) { i || (c ? c < 3 ? (c > 1 && (G.n = -1), d(c, u)) : G.n = u : G.v = u); try { if (f = 2, i) { if (c || (o = "next"), t = i[o]) { if (!(t = t.call(i, u))) throw TypeError("iterator result is not an object"); if (!t.done) return t; u = t.value, c < 2 && (c = 0); } else 1 === c && (t = i.return) && t.call(i), c < 2 && (u = TypeError("The iterator does not provide a '" + o + "' method"), c = 1); i = e; } else if ((t = (y = G.n < 0) ? u : r.call(n, G)) !== a) break; } catch (t) { i = e, c = 1, u = t; } finally { f = 1; } } return { value: t, done: y }; }; }(r, o, i), !0), u; } var a = {}; function Generator() {} function GeneratorFunction() {} function GeneratorFunctionPrototype() {} t = Object.getPrototypeOf; var c = [][n] ? t(t([][n]())) : (_regeneratorDefine2(t = {}, n, function () { return this; }), t), u = GeneratorFunctionPrototype.prototype = Generator.prototype = Object.create(c); function f(e) { return Object.setPrototypeOf ? Object.setPrototypeOf(e, GeneratorFunctionPrototype) : (e.__proto__ = GeneratorFunctionPrototype, _regeneratorDefine2(e, o, "GeneratorFunction")), e.prototype = Object.create(u), e; } return GeneratorFunction.prototype = GeneratorFunctionPrototype, _regeneratorDefine2(u, "constructor", GeneratorFunctionPrototype), _regeneratorDefine2(GeneratorFunctionPrototype, "constructor", GeneratorFunction), GeneratorFunction.displayName = "GeneratorFunction", _regeneratorDefine2(GeneratorFunctionPrototype, o, "GeneratorFunction"), _regeneratorDefine2(u), _regeneratorDefine2(u, o, "Generator"), _regeneratorDefine2(u, n, function () { return this; }), _regeneratorDefine2(u, "toString", function () { return "[object Generator]"; }), (_regenerator = function _regenerator() { return { w: i, m: f }; })(); }
function _regeneratorDefine2(e, r, n, t) { var i = Object.defineProperty; try { i({}, "", {}); } catch (e) { i = 0; } _regeneratorDefine2 = function _regeneratorDefine(e, r, n, t) { function o(r, n) { _regeneratorDefine2(e, r, function (e) { return this._invoke(r, n, e); }); } r ? i ? i(e, r, { value: n, enumerable: !t, configurable: !t, writable: !t }) : e[r] = n : (o("next", 0), o("throw", 1), o("return", 2)); }, _regeneratorDefine2(e, r, n, t); }
function asyncGeneratorStep(n, t, e, r, o, a, c) { try { var i = n[a](c), u = i.value; } catch (n) { return void e(n); } i.done ? t(u) : Promise.resolve(u).then(r, o); }
function _asyncToGenerator(n) { return function () { var t = this, e = arguments; return new Promise(function (r, o) { var a = n.apply(t, e); function _next(n) { asyncGeneratorStep(a, r, o, _next, _throw, "next", n); } function _throw(n) { asyncGeneratorStep(a, r, o, _next, _throw, "throw", n); } _next(void 0); }); }; }
/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

Office.onReady(function (info) {
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
function showMessage(message) {
  var isSuccess = arguments.length > 1 && arguments[1] !== undefined ? arguments[1] : true;
  var statusElement = document.getElementById("status-message");
  statusElement.innerText = message;
  statusElement.style.color = isSuccess ? "green" : "red";
}

// Affiche une confirmation personnalisée avant d'exécuter cleanAll
function showCleanAllConfirmation() {
  var statusElement = document.getElementById("status-message");
  statusElement.innerHTML = '<span style="color: orange;">⚠️ Êtes-vous sûr de vouloir tout supprimer ?</span><br>' + '<button id="confirm-clean-all-yes">Oui</button> ' + '<button id="confirm-clean-all-no">Non</button>';
  document.getElementById("confirm-clean-all-yes").onclick = /*#__PURE__*/_asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee() {
    return _regenerator().w(function (_context) {
      while (1) switch (_context.n) {
        case 0:
          statusElement.innerText = "Nettoyage en cours...";
          _context.n = 1;
          return cleanAll();
        case 1:
          return _context.a(2);
      }
    }, _callee);
  }));
  document.getElementById("confirm-clean-all-no").onclick = function () {
    statusElement.innerText = "Action annulée.";
    statusElement.style.color = "gray";
  };
}

// ========================================
// FONCTION 1 : Supprimer les images
// ========================================
function removeImages() {
  return _removeImages.apply(this, arguments);
} // Fonction utilitaire : calcule la taille estimée des images
function _removeImages() {
  _removeImages = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee2() {
    var item, _t;
    return _regenerator().w(function (_context2) {
      while (1) switch (_context2.p = _context2.n) {
        case 0:
          _context2.p = 0;
          item = Office.context.mailbox.item; // Vérifier si on peut modifier le corps
          if (!(!item.body || !item.body.setAsync)) {
            _context2.n = 1;
            break;
          }
          showMessage("❌ Modification du corps non disponible en mode lecture", false);
          return _context2.a(2);
        case 1:
          // Accédez au corps du mail
          item.body.getAsync(Office.CoercionType.Html, function (result) {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
              var htmlBody = result.value;

              // Extraire toutes les images et calculer la taille totale estimée
              var imgMatches = htmlBody.match(/<img[^>]*>/gi) || [];
              var totalSize = calculateImageSize(imgMatches);

              // Supprimer tous les tags <img>
              htmlBody = htmlBody.replace(/<img[^>]*>/gi, "");

              // Remplacer le corps du mail
              item.body.setAsync(htmlBody, {
                coercionType: Office.CoercionType.Html
              }, function (setResult) {
                if (setResult.status === Office.AsyncResultStatus.Succeeded) {
                  var sizeText = totalSize > 0 ? " (" + formatFileSize(totalSize) + ")" : "";
                  showMessage("✅ " + imgMatches.length + " image(s) supprimée(s) !" + sizeText, true);
                } else {
                  showMessage("❌ Erreur lors de la suppression des images: " + setResult.error.message, false);
                }
              });
            } else {
              showMessage("❌ Erreur de lecture du mail", false);
            }
          });
          _context2.n = 3;
          break;
        case 2:
          _context2.p = 2;
          _t = _context2.v;
          showMessage("❌ Erreur: " + _t.message, false);
        case 3:
          return _context2.a(2);
      }
    }, _callee2, null, [[0, 2]]);
  }));
  return _removeImages.apply(this, arguments);
}
function calculateImageSize(imgMatches) {
  var totalSize = 0;
  imgMatches.forEach(function (imgTag) {
    // Chercher l'attribut data-size si présent
    var dataSizeMatch = imgTag.match(/data-size="?(\d+)"?/i);
    if (dataSizeMatch) {
      totalSize += parseInt(dataSizeMatch[1], 10);
    } else {
      // Sinon, estimer basé sur width et height
      var widthMatch = imgTag.match(/width="?(\d+)"?/i);
      var heightMatch = imgTag.match(/height="?(\d+)"?/i);
      if (widthMatch && heightMatch) {
        var width = parseInt(widthMatch[1], 10);
        var height = parseInt(heightMatch[1], 10);
        // Estimation : ~5KB par 100x100 pixels
        var estimatedSize = Math.round(width * height / 2000) * 5120;
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
  var k = 1024;
  var sizeInKB = bytes / k;
  if (sizeInKB < 1) {
    return "< 1 KB économisé(s)";
  } else if (sizeInKB < 1024) {
    return Math.round(sizeInKB * 100) / 100 + " KB économisé(s)";
  } else if (sizeInKB < 1024 * 1024) {
    return Math.round(sizeInKB / 1024 * 100) / 100 + " MB économisé(s)";
  } else {
    return Math.round(sizeInKB / (1024 * 1024) * 100) / 100 + " GB économisé(s)";
  }
}

// ========================================
// FONCTION 2 : Supprimer les pièces jointes
// ========================================
function removeAttachments() {
  return _removeAttachments.apply(this, arguments);
} // ========================================
// FONCTION 3 : Garder les 2 derniers replies
// ========================================
function _removeAttachments() {
  _removeAttachments = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee3() {
    var item, callId, attachmentCount, totalSize, attachmentIds, successCount, _t2;
    return _regenerator().w(function (_context3) {
      while (1) switch (_context3.p = _context3.n) {
        case 0:
          _context3.p = 0;
          item = Office.context.mailbox.item;
          callId = Date.now() + Math.random(); // Identifiant unique pour cet appel
          // En mode composition, utiliser getAttachmentsAsync
          if (!(item.itemType === Office.MailboxEnums.ItemType.Message && typeof item.getAttachmentsAsync === 'function')) {
            _context3.n = 1;
            break;
          }
          item.getAttachmentsAsync(function (result) {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
              var attachments = result.value;
              if (!attachments || attachments.length === 0) {
                showMessage("ℹ️ Aucune pièce jointe à supprimer", true);
                return;
              }
              var attachmentCount = attachments.length;

              // Calculer la taille totale des pièces jointes
              var totalSize = 0;
              attachments.forEach(function (att) {
                if (att.size) {
                  totalSize += att.size;
                }
              });

              // Supprimer chaque pièce jointe
              var successCount = 0;
              attachments.forEach(function (attachment) {
                item.removeAttachmentAsync(attachment.id, function (removeResult) {
                  if (removeResult.status === Office.AsyncResultStatus.Succeeded) {
                    successCount++;
                    // Vérifier qu'on a bien le bon nombre et que c'est le bon appel
                    if (successCount === attachmentCount && attachmentCount > 0) {
                      var sizeText = totalSize > 0 ? " (" + formatFileSize(totalSize) + ")" : "";
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
          _context3.n = 3;
          break;
        case 1:
          if (!(!item.attachments || !Array.isArray(item.attachments) || item.attachments.length === 0)) {
            _context3.n = 2;
            break;
          }
          showMessage("ℹ️ Aucune pièce jointe à supprimer", true);
          return _context3.a(2);
        case 2:
          attachmentCount = item.attachments.length; // Calculer la taille totale des pièces jointes
          totalSize = 0;
          item.attachments.forEach(function (att) {
            if (att.size) {
              totalSize += att.size;
            }
          });
          attachmentIds = item.attachments.map(function (att) {
            return att.id;
          });
          successCount = 0;
          attachmentIds.forEach(function (attachmentId) {
            item.attachments.removeAttachmentAsync(attachmentId, function (result) {
              if (result.status === Office.AsyncResultStatus.Succeeded) {
                successCount++;
                // Vérifier qu'on a bien le bon nombre et que c'est le bon appel
                if (successCount === attachmentCount && attachmentCount > 0) {
                  var sizeText = totalSize > 0 ? " (" + formatFileSize(totalSize) + ")" : "";
                  showMessage("✅ " + successCount + " pièce(s) jointe(s) supprimée(s)!" + sizeText, true);
                }
              }
            });
          });
        case 3:
          _context3.n = 5;
          break;
        case 4:
          _context3.p = 4;
          _t2 = _context3.v;
          showMessage("❌ Erreur: " + _t2.message, false);
        case 5:
          return _context3.a(2);
      }
    }, _callee3, null, [[0, 4]]);
  }));
  return _removeAttachments.apply(this, arguments);
}
function keepTwoReplies() {
  return _keepTwoReplies.apply(this, arguments);
} // ========================================
// FONCTION 4 : Tout supprimer (réutilise les fonctions existantes)
// ========================================
function _keepTwoReplies() {
  _keepTwoReplies = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee4() {
    var item;
    return _regenerator().w(function (_context4) {
      while (1) switch (_context4.n) {
        case 0:
          try {
            item = Office.context.mailbox.item;
            item.body.getAsync(Office.CoercionType.Text, function (result) {
              if (result.status !== Office.AsyncResultStatus.Succeeded) {
                showMessage("❌ Erreur de lecture du mail", false);
                return;
              }
              var textBody = result.value || "";

              // Chercher tous les "De :" (qui marquent le début d'un reply)
              var deMatches = [];
              var deRegex = /De\s*:/gi;
              var match;
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
              var segments = [];
              for (var i = 0; i < deMatches.length; i++) {
                var start = deMatches[i];
                var end = i + 1 < deMatches.length ? deMatches[i + 1] : textBody.length;
                segments.push(textBody.substring(start, end));
              }

              // Le corps jusqu'au premier "De :"
              var firstReplyPos = deMatches[0];
              var cleanedText = textBody.substring(0, firstReplyPos);

              // Ajouter les 2 premiers segments (2 replies)
              cleanedText += segments[0] || "";
              if (segments.length > 1) {
                cleanedText += segments[1] || "";
              }

              // Remplacer le corps
              item.body.setAsync(cleanedText, {
                coercionType: Office.CoercionType.Text
              }, function (setResult) {
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
        case 1:
          return _context4.a(2);
      }
    }, _callee4);
  }));
  return _keepTwoReplies.apply(this, arguments);
}
function cleanAll() {
  return _cleanAll.apply(this, arguments);
}
function _cleanAll() {
  _cleanAll = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee5() {
    var _t3;
    return _regenerator().w(function (_context5) {
      while (1) switch (_context5.p = _context5.n) {
        case 0:
          _context5.p = 0;
          showMessage("Nettoyage en cours...", true);
          _context5.n = 1;
          return removeImages();
        case 1:
          _context5.n = 2;
          return removeAttachments();
        case 2:
          _context5.n = 3;
          return keepTwoReplies();
        case 3:
          showMessage("✅ Mail complètement nettoyé!", true);
          _context5.n = 5;
          break;
        case 4:
          _context5.p = 4;
          _t3 = _context5.v;
          showMessage("❌ Erreur: " + _t3.message, false);
        case 5:
          return _context5.a(2);
      }
    }, _callee5, null, [[0, 4]]);
  }));
  return _cleanAll.apply(this, arguments);
}
}();
// This entry needs to be wrapped in an IIFE because it needs to be in strict mode.
!function() {
"use strict";
/*!************************************!*\
  !*** ./src/taskpane/taskpane.html ***!
  \************************************/
__webpack_require__.r(__webpack_exports__);
// Imports
var ___HTML_LOADER_IMPORT_0___ = new URL(/* asset import */ __webpack_require__(/*! ./taskpane.css */ "./src/taskpane/taskpane.css"), __webpack_require__.b);
// Module
var code = "<!DOCTYPE html>\n<html>\n\n<head>\n    <meta charset=\"UTF-8\" />\n    <meta http-equiv=\"X-UA-Compatible\" content=\"IE=Edge\" />\n    <meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">\n    <title>Clean It Now Add-in</title>\n\n    <!-- Office JavaScript API -->\n    <" + "script type=\"text/javascript\" src=\"https://appsforoffice.microsoft.com/lib/1/hosted/office.js\"><" + "/script>\n\n    <!-- For more information on Fluent UI, visit https://developer.microsoft.com/fluentui#/. -->\n    <link rel=\"stylesheet\" href=\"https://res-1.cdn.office.net/files/fabric-cdn-prod_20230815.002/office-ui-fabric-core/11.1.0/css/fabric.min.css\"/>\n\n    <!-- Template styles -->\n    <link href=\"" + ___HTML_LOADER_IMPORT_0___ + "\" rel=\"stylesheet\" type=\"text/css\" />\n</head>\n\n<body class=\"ms-font-m ms-welcome ms-Fabric\">\n    <header class=\"ms-welcome__header ms-bgColor-neutralLighter\">\n        <h1 class=\"ms-font-su\">Welcome</h1>\n    </header>\n    <section id=\"sideload-msg\" class=\"ms-welcome__main\">\n        <h2 class=\"ms-font-xl\">Please <a target=\"_blank\" href=\"https://learn.microsoft.com/office/dev/add-ins/testing/test-debug-office-add-ins#sideload-an-office-add-in-for-testing\">sideload</a> your add-in to see app body.</h2>\n    </section>\n    <main id=\"app-body\" class=\"ms-welcome__main\" style=\"display: none;\">\n                <p class=\"ms-font-l\"><b>Nettoyage du mail</b></p>\n        <div role=\"button\" id=\"remove-images\" class=\"ms-welcome__action ms-Button ms-Button--hero ms-font-xl\">\n            <span class=\"ms-Button-label\">🖼️ Supprimer les images</span>\n        </div>\n        <div role=\"button\" id=\"keep-two-replies\" class=\"ms-welcome__action ms-Button ms-Button--hero ms-font-xl\">\n            <span class=\"ms-Button-label\">💬 Garder les 2 derniers replies</span>\n        </div>\n        <div role=\"button\" id=\"remove-attachments\" class=\"ms-welcome__action ms-Button ms-Button--hero ms-font-xl\">\n            <span class=\"ms-Button-label\">📎 Supprimer les PJ</span>\n        </div>\n        <div role=\"button\" id=\"clean-all\" class=\"ms-welcome__action ms-Button ms-Button--hero ms-font-xl\">\n            <span class=\"ms-Button-label\">🧹 Tout supprimer</span>\n        </div>\n        <p><label id=\"status-message\"></label></p>\n    </main>\n</body>\n\n</html>\n";
// Exports
/* harmony default export */ __webpack_exports__["default"] = (code);
}();
/******/ })()
;
//# sourceMappingURL=taskpane.js.map