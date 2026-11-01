// Outils du test de ModGraphMail : transcription VBA → VBScript, équivalents de test, faux Microsoft Graph.
import http from 'node:http';

// ---------- Transcription VBA → VBScript ----------
export function versVbscript(vba) {
  let t = vba.replace(/\r\n/g, '\n');
  t = t.replace(/^Attribute .*\n/gm, '');
  t = t.replace(/'#coffre-debut[\s\S]*?'#coffre-fin\n/, '').replace(/'#binaire-debut[\s\S]*?'#binaire-fin\n?/, '');
  const sansTypes = (s) => s.replace(/\s+As\s+(?:New\s+)?[\w.]+/g, '');
  return t.split('\n').map((l) => {
    const m = /^(\s*)(Public |Private )?Const (\w+) As \w+ = (.*)$/.exec(l);
    if (m) return `${m[1]}Const ${m[3]} = ${m[4]}`;
    if (/^(Private|Public) m\w+ As /.test(l)) return l.replace(/^(Private|Public) (\w+) As \w+/, 'Dim $2');
    if (/^\s*Dim /.test(l)) return sansTypes(l).replace(/(\w+)\(\)/g, '$1');
    if (/^\s*ReDim /.test(l)) return l.replace(/\(0 To /g, '(');
    const sig = /^((?:Public |Private )?(?:Function|Sub) \w+)\((.*)\)(?:\s+As\s+[\w()]+)?\s*$/.exec(l);
    if (sig) {
      const params = sig[2].split(',').map((p) => sansTypes(p.replace(/\bOptional\s+/, '').replace(/\s*=\s*.*$/, '')).trim()).filter(Boolean);
      return `${sig[1]}(${params.join(', ')})`;
    }
    return l;
  }).join('\r\n').replace(/\bAscW\(/g, 'TAscW(').replace(/\bChrW\(/g, 'TChrW(');
}

// Le VBScript de Wine 9 n'a ni AscW ni ChrW : équivalents limités aux caractères des tests (page 1252).
export const UNICODE = `
Function TAscW(ByVal ch)
    Dim c: c = Asc(ch)
    If c = 151 Then c = 8212
    If c = 128 Then c = 8364
    TAscW = c
End Function
Function TChrW(ByVal c)
    If c < 0 Then c = c + 65536
    If c = 8212 Then c = 151
    If c = 8364 Then c = 128
    TChrW = Chr(c)
End Function
`;
// Sortie console de cscript : page OEM 850.
const CP850 = { 0x82: 'é', 0x8a: 'è', 0x85: 'à', 0x88: 'ê', 0x87: 'ç', 0xae: '«', 0xaf: '»', 0x97: 'ù', 0x8c: 'î', 0x93: 'ô', 0x90: 'É', 0xff: '\u00a0' };
export const oem = (b) => [...b].map((x) => (x < 128 ? String.fromCharCode(x) : CP850[x] ?? '?')).join('');

// Équivalents de test du bloc « binaire » : fichiers texte ASCII, Base64 calculé en VBScript.
export const STUBS = `
Function FichierExiste(ByVal chemin)
    FichierExiste = CreateObject("Scripting.FileSystemObject").FileExists(chemin)
End Function
Function TailleFichier(ByVal chemin)
    TailleFichier = CreateObject("Scripting.FileSystemObject").GetFile(chemin).Size
End Function
Function LireOctets(ByVal chemin)
    Dim f: Set f = CreateObject("Scripting.FileSystemObject").OpenTextFile(chemin, 1)
    If f.AtEndOfStream Then LireOctets = "" Else LireOctets = f.ReadAll
    f.Close
End Function
Function Tranche(ByRef octets, ByVal debut, ByVal n)
    Tranche = Mid(octets, debut + 1, n)
End Function
Function Base64Fichier(ByVal chemin)
    Const T = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
    Dim s, i, a, b, c, r: s = LireOctets(chemin)
    For i = 1 To Len(s) Step 3
        a = Asc(Mid(s, i, 1)): b = -1: c = -1
        If i + 1 <= Len(s) Then b = Asc(Mid(s, i + 1, 1))
        If i + 2 <= Len(s) Then c = Asc(Mid(s, i + 2, 1))
        r = r & Mid(T, (a \\ 4) + 1, 1)
        If b < 0 Then
            r = r & Mid(T, ((a And 3) * 16) + 1, 1) & "=="
        Else
            r = r & Mid(T, ((a And 3) * 16 + (b \\ 16)) + 1, 1)
            If c < 0 Then
                r = r & Mid(T, ((b And 15) * 4) + 1, 1) & "="
            Else
                r = r & Mid(T, ((b And 15) * 4 + (c \\ 64)) + 1, 1) & Mid(T, (c And 63) + 1, 1)
            End If
        End If
    Next
    Base64Fichier = r
End Function
Sub LireCoffre(ByVal cible, ByRef utilisateur, ByRef secret)
End Sub
`;

// ---------- Faux Graph ----------
export function fauxGraph() {
  const journal = [];
  const televersements = new Map();
  let n = 0;
  const serveur = http.createServer((req, res) => {
    const morceaux = [];
    req.on('data', (c) => morceaux.push(c));
    req.on('end', () => {
      const corps = Buffer.concat(morceaux);
      const url = new URL(req.url, 'http://x');
      const entree = { methode: req.method, chemin: decodeURIComponent(url.pathname), entetes: req.headers, corps };
      journal.push(entree);
      const json = (statut, o) => { res.writeHead(statut, { 'content-type': 'application/json' }); res.end(o ? JSON.stringify(o) : ''); };
      const p = entree.chemin;
      if (p === '/connexion/exemple.onmicrosoft.com/oauth2/v2.0/token') {
        const f = new URLSearchParams(corps.toString());
        if (f.get('client_secret') !== 's3cr~t.x_y+/=') return json(401, { error: 'invalid_client', error_description: 'AADSTS7000215: Invalid client secret provided.\r\nTrace ID: 1' });
        return json(200, { token_type: 'Bearer', expires_in: 3599, access_token: 'jeton-1' });
      }
      if (req.headers.authorization !== 'Bearer jeton-1' && !p.startsWith('/televersement/')) return json(401, { error: { code: 'InvalidAuthenticationToken', message: 'jeton absent' } });
      if (p === '/v1.0/users/interdit@exemple.invalid/sendMail') return json(403, { error: { code: 'ErrorAccessDenied', message: 'Access is denied. Check credentials and try again.' } });
      if (/^\/v1\.0\/users\/[^/]+\/sendMail$/.test(p)) return json(202);
      if (/^\/v1\.0\/users\/[^/]+\/messages$/.test(p)) { n++; return json(201, { '@odata.etag': 'W/"x"', id: `AAMk${n}+/=`, webLink: `https://outlook.office365.com/owa/?ItemID=AAMk${n}&exvsurl=1&viewmodel=ReadMessageItem` }); }
      if (/\/attachments\/createUploadSession$/.test(p)) {
        const id = `s${journal.length}`;
        televersements.set(id, { attendu: JSON.parse(corps).AttachmentItem.size, recu: [] });
        return json(201, { uploadUrl: `http://127.0.0.1:${serveur.address().port}/televersement/${id}?sig=a%2Fb&x=1`, expirationDateTime: '2026-10-24T12:00:00Z' });
      }
      if (/\/attachments$/.test(p)) return json(201, { id: 'piece' });
      if (/\/send$/.test(p)) return json(202);
      const t = /^\/televersement\/(\w+)$/.exec(p);
      if (t && req.method === 'PUT') {
        const s = televersements.get(t[1]);
        s.recu.push({ plage: req.headers['content-range'], corps });
        const fin = s.recu.reduce((x, r) => x + r.corps.length, 0) === s.attendu;
        return json(fin ? 201 : 200, fin ? null : { nextExpectedRanges: ['…'] });
      }
      json(404, { error: { code: 'ErrorItemNotFound', message: p } });
    });
  });
  return new Promise((ok) => serveur.listen(0, '127.0.0.1', () => ok({ serveur, journal, televersements, port: serveur.address().port })));
}

