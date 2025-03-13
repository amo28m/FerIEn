// Konfiguration für MSAL (Microsoft Authentication Library)
const msalConfig = {
  auth: {
    clientId: 'f4602006-b304-4530-8e4e-7c31c9b3cb2e', // Die Client-ID Ihrer Anwendung
    authority: 'https://login.microsoftonline.com/2356b269-1a6e-4033-a730-46e40484e6b5', // Die Autorität (Azure AD Tenant)
    redirectUri: 'https://amo28m.github.io/FerIEn/src/taskpane/taskpane.html', // Die URI, zu der nach der Authentifizierung umgeleitet wird
  },
  cache: {
    cacheLocation: 'localStorage', // Wo die Sitzungsdaten gespeichert werden
    storeAuthStateInCookie: true, // Authentifizierungsstatus auch in Cookies speichern
  },
};

// Anmeldeanforderung mit den benötigten Berechtigungen
const loginRequest = {
  scopes: ['Calendars.ReadWrite', 'User.Read'],
};

// Globale Variablen
let msalInstance;
let projectCount = 1; // Startwert für die Anzahl der Projekte
const additionalEmail = 'gz.ma-abwesenheiten@ie-group.com'; // Zusätzliche E-Mail-Adresse

// Event-Listener, der ausgeführt wird, sobald das DOM vollständig geladen ist
document.addEventListener('DOMContentLoaded', function () {
  // Initialisiert die MSAL-Instanz
  msalInstance = new msal.PublicClientApplication(msalConfig);

  // Wartet, bis Office bereit ist
  Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
      const holidayForm = document.getElementById('holidayForm');
      // Fügt Event-Handler für das Formular und die Buttons hinzu
      holidayForm.onsubmit = submitHoliday;
      document.getElementById('addProjectButton').onclick = addProjectFields;
      document.getElementById('removeProjectButton').onclick = removeProjectFields;
      addProjectFields(); // Fügt initial ein Projektfeld hinzu

      // Nach kurzer Verzögerung: Erzwinge ein Neurendern des Formulars, um Fokusprobleme bei den statischen Feldern zu beheben
      setTimeout(() => {
        const formHtml = holidayForm.innerHTML;
        holidayForm.innerHTML = formHtml;
        // Hänge die Event-Handler erneut an
        holidayForm.onsubmit = submitHoliday;
        document.getElementById('addProjectButton').onclick = addProjectFields;
        document.getElementById('removeProjectButton').onclick = removeProjectFields;
      }, 500);
    }
  });
});

// Funktion zum Hinzufügen von Projektfeldern
function addProjectFields() {
  projectCount++; // Erhöht die Projektanzahl

  // Erstellt eine neue Projektgruppe im DOM
  const projectGroup = document.createElement('div');
  projectGroup.className = 'project-group';
  projectGroup.id = `projectGroup${projectCount}`;
  projectGroup.innerHTML = `
    <hr class="divider">
    <div class="form-group">
      <label for="projectNumber${projectCount}">Projektnummer/Funktion:</label>
      <input type="text" id="projectNumber${projectCount}" required placeholder="Projekt/Funktion ...">
    </div>
    <div class="form-group">
      <label for="projectManager${projectCount}">Projektleiter:</label>
      <input type="text" id="projectManager${projectCount}" required placeholder="email@ie-group.com, ...">
    </div>
    <div class="form-group">
      <label for="projectDeputy${projectCount}">Stellvertreter des Projektes:</label>
      <input type="text" id="projectDeputy${projectCount}" required placeholder="email@ie-group.com, ...">
    </div>
  `;
  document.getElementById('additionalProjects').appendChild(projectGroup); // Fügt die Projektgruppe dem DOM hinzu
}

// Funktion zum Entfernen von Projektfeldern
function removeProjectFields() {
  if (projectCount > 0) {
    const projectGroup = document.getElementById(`projectGroup${projectCount}`);
    if (projectGroup) {
      projectGroup.remove(); // Entfernt die letzte Projektgruppe
      projectCount--;
    }
  } else {
    showConfirmationMessage('Es gibt keine Projekte mehr zum Entfernen.', true); // Meldung anzeigen, wenn keine Projekte mehr vorhanden sind
  }
}

// Funktion zum Verarbeiten des Urlaubsformulars
function submitHoliday(event) {
  event.preventDefault(); // Verhindert das Standardverhalten des Formulars

  // Liest die Eingabewerte aus den Formularfeldern
  const startDate = document.getElementById('startDate').value;
  let endDate = document.getElementById('endDate').value;
  endDate = new Date(endDate);
  endDate.setHours(23, 59, 59); // Setzt die Zeit auf Ende des Tages

  // Formatiert das Enddatum korrekt
  const localEndDate = endDate.getFullYear() + '-' +
    String(endDate.getMonth() + 1).padStart(2, '0') + '-' +
    String(endDate.getDate()).padStart(2, '0') + 'T' +
    String(endDate.getHours()).padStart(2, '0') + ':' +
    String(endDate.getMinutes()).padStart(2, '0') + ':' +
    String(endDate.getSeconds()).padStart(2, '0');

  const reason = document.getElementById('reason').value;
  const deputy = document.getElementById('deputy').value;

  // Sammelt alle Projektinformationen
  const projectFields = [];
  for (let i = 1; i <= projectCount; i++) {
    const projectNumber = document.getElementById(`projectNumber${i}`);
    const projectManager = document.getElementById(`projectManager${i}`);
    const projectDeputy = document.getElementById(`projectDeputy${i}`);

    if (projectNumber && projectManager && projectDeputy) {
      projectFields.push({
        number: projectNumber.value,
        manager: projectManager.value,
        deputy: projectDeputy.value,
      });
    }
  }

  // Überprüft, ob alle erforderlichen Felder ausgefüllt sind
  if (
    startDate &&
    endDate &&
    reason &&
    deputy &&
    projectFields.every((field) => field.number && field.manager && field.deputy)
  ) {
    // Überprüfen Sie die E-Mail-Adressen des Stellvertreters
    const deputyEmails = parseEmails(deputy);
    if (deputyEmails.length === 0) {
      showConfirmationMessage('Bitte geben Sie mindestens eine gültige @ie-group.com E-Mail-Adresse für den Stellvertreter an.', true);
      return;
    }

    // Überprüfen Sie die E-Mail-Adressen in den Projektfeldern
    for (const field of projectFields) {
      const managerEmails = parseEmails(field.manager);
      const deputyEmails = parseEmails(field.deputy);

      if (managerEmails.length === 0 || deputyEmails.length === 0) {
        showConfirmationMessage('Bitte geben Sie gültige @ie-group.com E-Mail-Adressen für Projektleiter und Stellvertreter an.', true);
        return;
      }
    }

    resetForm(); // Setzt das Formular zurück

    // Startet den Anmeldeprozess
    msalInstance
      .loginPopup(loginRequest)
      .then((loginResponse) => {
        const account = msalInstance.getAccountByHomeId(loginResponse.account.homeAccountId);
        const accessTokenRequest = {
          scopes: ['Calendars.ReadWrite', 'User.Read'],
          account: account,
        };

        // Fordert ein Zugriffstoken an
        msalInstance
          .acquireTokenSilent(accessTokenRequest)
          .then((tokenResponse) => {
            const accessToken = tokenResponse.accessToken;

            getUserName(accessToken) // Ruft den Benutzernamen ab
              .then((senderName) => {
                const subject = `${senderName}: ${reason}`; // Erstellt den Betreff
                const bodyContent = generateBodyContent(startDate, localEndDate, reason, deputy, projectFields); // Generiert den Nachrichtentext

                // Sammelt alle E-Mail-Adressen der Teilnehmer
                const allAttendees = deputyEmails.concat(
                  ...projectFields.map((field) => parseEmails(field.manager)),
                  ...projectFields.map((field) => parseEmails(field.deputy)),
                  additionalEmail
                );

                // Erstellt das Ereignis im Kalender
                createEvent(startDate, localEndDate, subject, bodyContent, Office.context.mailbox.userProfile.emailAddress, allAttendees, accessToken, 'free')
                  .then((eventId) => {
                    // Aktualisiert den Status des Ereignisses auf 'beschäftigt'
                    updateEventStatus(eventId, 'busy', accessToken)
                      .then(() => {
                        showConfirmationMessage('Urlaub erfolgreich eingetragen!');
                      })
                      .catch((error) => {
                        console.error('Fehler beim Aktualisieren des Ereignisses:', error);
                        showConfirmationMessage('Fehler beim Aktualisieren des Ereignisses.', true);
                      });
                  })
                  .catch((error) => {
                    console.error('Fehler beim Erstellen des Ereignisses:', error);
                    showConfirmationMessage('Fehler beim Erstellen des Ereignisses.', true);
                  });
              })
              .catch((error) => {
                console.error('Fehler beim Abrufen des Benutzernamens:', error);
                showConfirmationMessage('Fehler beim Abrufen des Benutzernamens.', true);
              });
          })
          .catch((error) => {
            console.error('Fehler beim Abrufen des Zugriffstokens:', error);
            showConfirmationMessage('Fehler beim Abrufen des Zugriffstokens.', true);
          });
      })
      .catch((error) => {
        console.error('Fehler bei der Anmeldung:', error);
        showConfirmationMessage('Fehler bei der Anmeldung.', true);
      });
  } else {
    showConfirmationMessage('Bitte alle Felder ausfüllen.', true); // Meldung anzeigen, wenn Felder fehlen
  }
}

// Validiert eine E-Mail-Adresse, die auf @ie-group.com endet
function isValidEmail(email) {
  const emailRegex = /^[^\s@]+@ie-group\.com$/i;
  return emailRegex.test(email);
}

// Funktion zum Parsen von E-Mail-Adressen aus einem String
function parseEmails(emailString) {
  const emails = emailString
    .split(',')
    .map((email) => email.trim())
    .filter((email) => email.length > 0);

  const invalidEmails = emails.filter((email) => !isValidEmail(email));

  if (invalidEmails.length > 0) {
    showConfirmationMessage(`Ungültige E-Mail-Adresse(n): ${invalidEmails.join(', ')}. Bitte verwenden Sie nur @ie-group.com-Adressen.`, true);
    return [];
  }

  return emails;
}

// Erstellt ein neues Kalenderereignis
function createEvent(
  startDate,
  endDateTime,
  subject,
  bodyContent,
  organizerEmail,
  attendeesEmails,
  accessToken,
  showAs
) {
  const attendees = attendeesEmails.map((email) => ({
    emailAddress: {
      address: email,
    },
    type: 'required',
  }));

  const event = {
    subject: subject,
    start: {
      dateTime: `${startDate}T00:00:00`,
      timeZone: 'Europe/Zurich',
    },
    end: {
      dateTime: endDateTime,
      timeZone: 'Europe/Zurich',
    },
    body: {
      contentType: 'HTML',
      content: bodyContent,
    },
    showAs: showAs,
    attendees: attendees,
  };

  // Sendet eine Anfrage an die Microsoft Graph API zum Erstellen des Ereignisses
  return fetch('https://graph.microsoft.com/v1.0/me/events', {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${accessToken}`,
      'Content-Type': 'application/json',
    },
    body: JSON.stringify(event),
  }).then((response) => {
    if (!response.ok) {
      return response.json().then((error) => {
        throw new Error(`Fehler beim Erstellen des Ereignisses für ${organizerEmail}: ${error.message}`);
      });
    }
    return response.json().then((event) => event.id); // Gibt die Ereignis-ID zurück
  });
}

// Aktualisiert den Status eines bestehenden Ereignisses
function updateEventStatus(eventId, showAs, accessToken) {
  const update = {
    showAs: showAs,
  };

  // Sendet eine PATCH-Anfrage an die Microsoft Graph API
  return fetch(`https://graph.microsoft.com/v1.0/me/events/${eventId}`, {
    method: 'PATCH',
    headers: {
      Authorization: `Bearer ${accessToken}`,
      'Content-Type': 'application/json',
    },
    body: JSON.stringify(update),
  }).then((response) => {
    if (!response.ok) {
      return response.json().then((error) => {
        throw new Error(`Fehler beim Aktualisieren des Ereignisses: ${error.message}`);
      });
    }
  });
}

// Setzt das Formular zurück
function resetForm() {
  const startDateField = document.getElementById('startDate');
  const endDateField = document.getElementById('endDate');
  const reasonField = document.getElementById('reason');
  const deputyField = document.getElementById('deputy');
  const projectNumberField = document.getElementById('projectNumber1');
  const projectManagerField = document.getElementById('projectManager1');
  const projectDeputyField = document.getElementById('projectDeputy1');

  if (startDateField) startDateField.value = '';
  if (endDateField) endDateField.value = '';
  if (reasonField) reasonField.value = '';
  if (deputyField) deputyField.value = '';
  if (projectNumberField) projectNumberField.value = '';
  if (projectManagerField) projectManagerField.value = '';
  if (projectDeputyField) projectDeputyField.value = '';

  document.getElementById('additionalProjects').innerHTML = ''; // Entfernt zusätzliche Projekte
  projectCount = 1; // Setzt die Projektanzahl zurück
}

// Zeigt eine Bestätigungsmeldung an
function showConfirmationMessage(message, isError = false) {
  const confirmationMessage = document.getElementById('confirmationMessage');
  confirmationMessage.innerText = message;
  confirmationMessage.style.display = 'block';

  if (isError) {
    confirmationMessage.classList.add('error');
  } else {
    confirmationMessage.classList.remove('error');
  }
}

// Generiert den Inhalt für den Nachrichtentext
function generateBodyContent(startDate, endDate, reason, deputy, projectFields) {
  let content = `<div style="font-family: Arial; font-size: 10pt;">
                      Ferienabwesenheit von ${formatDate(startDate)} bis ${formatDate(endDate)}.<br>
                      Vorgesetzter: ${deputy}<br>
                      Grund: ${reason}<br>`;

  projectFields.forEach((field, index) => {
    content += `Projektnummer ${index + 1}: ${field.number}, Projektleiter: ${field.manager}, Projektstellvertreter: ${field.deputy}<br>`;
  });

  content += '</div>';
  return content;
}

// Formatiert ein Datum im Format TT.MM.JJJJ
function formatDate(dateString) {
  const date = new Date(dateString);
  const day = String(date.getDate()).padStart(2, '0');
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const year = date.getFullYear();
  return `${day}.${month}.${year}`;
}

// Ruft den Benutzernamen über die Microsoft Graph API ab
function getUserName(accessToken) {
  return fetch('https://graph.microsoft.com/v1.0/me', {
    method: 'GET',
    headers: {
      Authorization: `Bearer ${accessToken}`,
      'Content-Type': 'application/json',
    },
  }).then((response) => {
    if (!response.ok) {
      return response.json().then((error) => {
        throw new Error(`Fehler beim Abrufen des Benutzernamens: ${error.message}`);
      });
    }
    return response.json().then((user) => user.displayName);
  });
}
