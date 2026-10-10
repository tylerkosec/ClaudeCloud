#target photoshop
// LL Johnson pitch deck: Photoshop template builder (Tyler Kosec Media)
// Run in Photoshop: File > Scripts > Browse... > pick this file.
// Builds 10 slides (1920x1080) as layered PSDs with editable text, editable color
// fill layers and photo frames. Your stills go in as smart objects clipped to each frame.
// To swap a photo later: right-click its layer > Replace Contents.

app.bringToFront();

(function () {
    var W = 1920, H = 1080;

    // ---- palette ----
    var DARK = "16261D", DARK2 = "22362A", LIGHT = "F4F1E8", LIGHT2 = "E8E3D5";
    var ACCENT = "E08A2E", ACCENT_DK = "9A4F12", BODY_L = "4A5249", BODY_D = "C8D2C6";
    var BROWN = "3A2A12", LINE_D = "3E5446", FRAME = "5A6B60";

    var MISSING = {};
    var FONTMAP = null;
    var stills = [];
    var nextStill = 0;
    var placedCount = 0, frameCount = 0;
    var savedFiles = [];

    // ---- preferences ----
    var oldRuler = app.preferences.rulerUnits;
    var oldType = app.preferences.typeUnits;
    var oldDialogs = app.displayDialogs;

    function restore() {
        app.preferences.rulerUnits = oldRuler;
        app.preferences.typeUnits = oldType;
        app.displayDialogs = oldDialogs;
    }

    // ---- helpers ----
    function cid(s) { return charIDToTypeID(s); }
    function sid(s) { return stringIDToTypeID(s); }

    function rgb(hex) {
        var c = new SolidColor();
        c.rgb.hexValue = hex;
        return c;
    }

    function hexParts(hex) {
        return [parseInt(hex.substr(0, 2), 16), parseInt(hex.substr(2, 2), 16), parseInt(hex.substr(4, 2), 16)];
    }

    function rectPts(x, y, w, h) {
        return [[x, y], [x + w, y], [x + w, y + h], [x, y + h]];
    }

    function unclip(layer) {
        try { if (layer.grouped) { layer.grouped = false; } } catch (e) {}
    }

    function findFont(family, styles, fallback) {
        if (!FONTMAP) {
            FONTMAP = {};
            for (var i = 0; i < app.fonts.length; i++) {
                var f = app.fonts[i];
                FONTMAP[(f.family + "|" + f.style).toLowerCase()] = f.postScriptName;
            }
        }
        for (var j = 0; j < styles.length; j++) {
            var key = (family + "|" + styles[j]).toLowerCase();
            if (FONTMAP[key]) { return FONTMAP[key]; }
        }
        MISSING[family + " " + styles[0]] = true;
        return fallback;
    }

    var F = {};
    function loadFonts() {
        F.OSW = findFont("Oswald", ["SemiBold", "Semibold", "Bold", "Medium"], "Arial-BoldMT");
        F.SANS = findFont("Public Sans", ["Regular"], "ArialMT");
        F.SANSSB = findFont("Public Sans", ["SemiBold", "Semibold", "Bold"], "Arial-BoldMT");
        F.SANSB = findFont("Public Sans", ["Bold"], "Arial-BoldMT");
    }

    // Solid color fill layer. If a selection is active it becomes the layer's mask.
    function solidFill(hex) {
        var p = hexParts(hex);
        var desc = new ActionDescriptor();
        var ref = new ActionReference();
        ref.putClass(sid("contentLayer"));
        desc.putReference(cid("null"), ref);
        var layerDesc = new ActionDescriptor();
        var colorDesc = new ActionDescriptor();
        var c = new ActionDescriptor();
        c.putDouble(cid("Rd  "), p[0]);
        c.putDouble(cid("Grn "), p[1]);
        c.putDouble(cid("Bl  "), p[2]);
        colorDesc.putObject(cid("Clr "), cid("RGBC"), c);
        layerDesc.putObject(cid("Type"), sid("solidColorLayer"), colorDesc);
        desc.putObject(cid("Usng"), sid("contentLayer"), layerDesc);
        executeAction(cid("Mk  "), desc, DialogModes.NO);
        return app.activeDocument.activeLayer;
    }

    function bg(hex) {
        var d = app.activeDocument;
        d.selection.deselect();
        var l = solidFill(hex);
        l.name = "Background color (double-click to change)";
        return l;
    }

    function rect(name, x, y, w, h, hex, opacity, hidden) {
        var d = app.activeDocument;
        d.selection.select(rectPts(x, y, w, h));
        var l = solidFill(hex);
        d.selection.deselect();
        unclip(l);
        l.name = name;
        if (opacity !== undefined && opacity !== null) { l.opacity = opacity; }
        if (hidden) { l.visible = false; }
        return l;
    }

    // Full-canvas linear gradient fill layer, left to right, in one color with opacity stops.
    // stops: [[locationPercent, opacityPercent], ...]
    function scrim(name, hex, stops) {
        try {
            return gradientScrim(name, hex, stops);
        } catch (e) {
            // Fallback: a flat overlay at the average opacity, still editable.
            var l = solidFill(hex);
            unclip(l);
            l.name = name + " (flat fallback)";
            l.opacity = 60;
            return l;
        }
    }

    function gradientScrim(name, hex, stops) {
        var p = hexParts(hex);
        var d = app.activeDocument;
        d.selection.deselect();
        var desc = new ActionDescriptor();
        var ref = new ActionReference();
        ref.putClass(sid("contentLayer"));
        desc.putReference(cid("null"), ref);
        var layerDesc = new ActionDescriptor();
        var gradLayer = new ActionDescriptor();
        gradLayer.putUnitDouble(cid("Angl"), cid("#Ang"), 0);
        gradLayer.putEnumerated(cid("Type"), cid("GrdT"), cid("Lnr "));
        var grad = new ActionDescriptor();
        grad.putString(cid("Nm  "), name);
        grad.putEnumerated(cid("GrdF"), cid("GrdF"), cid("CstS"));
        grad.putDouble(cid("Intr"), 4096);
        var colors = new ActionList();
        var ends = [0, 100];
        for (var i = 0; i < ends.length; i++) {
            var cs = new ActionDescriptor();
            var c = new ActionDescriptor();
            c.putDouble(cid("Rd  "), p[0]);
            c.putDouble(cid("Grn "), p[1]);
            c.putDouble(cid("Bl  "), p[2]);
            cs.putObject(cid("Clr "), cid("RGBC"), c);
            cs.putEnumerated(cid("Type"), cid("Clry"), cid("UsrS"));
            cs.putInteger(cid("Lctn"), Math.round(ends[i] * 40.96));
            cs.putInteger(cid("Mdpn"), 50);
            colors.putObject(cid("Clrt"), cs);
        }
        grad.putList(cid("Clrs"), colors);
        var trans = new ActionList();
        for (var k = 0; k < stops.length; k++) {
            var ts = new ActionDescriptor();
            ts.putUnitDouble(cid("Opct"), cid("#Prc"), stops[k][1]);
            ts.putInteger(cid("Lctn"), Math.round(stops[k][0] * 40.96));
            ts.putInteger(cid("Mdpn"), 50);
            trans.putObject(cid("TrnS"), ts);
        }
        grad.putList(cid("Trns"), trans);
        gradLayer.putObject(cid("Grad"), cid("Grdn"), grad);
        layerDesc.putObject(cid("Type"), sid("gradientLayer"), gradLayer);
        desc.putObject(cid("Usng"), sid("contentLayer"), layerDesc);
        executeAction(cid("Mk  "), desc, DialogModes.NO);
        var l = app.activeDocument.activeLayer;
        unclip(l);
        l.name = name;
        return l;
    }

    function placeFile(file) {
        var desc = new ActionDescriptor();
        desc.putPath(cid("null"), file);
        desc.putEnumerated(cid("FTcs"), cid("QCSt"), cid("Qcsa"));
        var off = new ActionDescriptor();
        off.putUnitDouble(cid("Hrzn"), cid("#Pxl"), 0);
        off.putUnitDouble(cid("Vrtc"), cid("#Pxl"), 0);
        desc.putObject(cid("Ofst"), cid("Ofst"), off);
        executeAction(cid("Plc "), desc, DialogModes.NO);
        return app.activeDocument.activeLayer;
    }

    function px(v) { return v.as("px"); }

    // Photo frame: a gray pixel layer, with the next still placed above it as a
    // smart object, scaled to cover the frame and clipped to it.
    function photo(name, x, y, w, h) {
        var d = app.activeDocument;
        frameCount++;
        var base = d.artLayers.add();
        unclip(base);
        base.name = name + " - FRAME (keep under the photo)";
        d.selection.select(rectPts(x, y, w, h));
        d.selection.fill(rgb(FRAME));
        d.selection.deselect();

        var file = null;
        if (nextStill < stills.length) { file = stills[nextStill]; nextStill++; }
        if (file) {
            try {
                var l = placeFile(file);
                var b = l.bounds;
                var lw = px(b[2]) - px(b[0]), lh = px(b[3]) - px(b[1]);
                var s = Math.max(w / lw, h / lh) * 100;
                l.resize(s, s, AnchorPosition.MIDDLECENTER);
                b = l.bounds;
                var cx = (px(b[0]) + px(b[2])) / 2, cy = (px(b[1]) + px(b[3])) / 2;
                l.translate(x + w / 2 - cx, y + h / 2 - cy);
                l.grouped = true;
                l.name = name + " - PHOTO: " + decodeURI(file.name) + " (right-click > Replace Contents)";
                placedCount++;
                return l;
            } catch (e) {
                // fall through to a placeholder label
            }
        }
        var t = text(name + " - empty frame label (delete after placing a still)",
            "DROP A STILL HERE\rFile > Place Embedded, then Alt-click between layers to clip",
            x + 40, y + h / 2 - 40, w - 80, 120,
            { f: "SANSB", s: 28, ld: 38, c: LIGHT, j: "C" });
        return t;
    }

    // Paragraph text layer. o: f (font key), s (size px), ld (leading px), c (hex),
    // tr (tracking), caps, j ("R" right or "C" center), hidden
    function text(name, contents, x, y, w, h, o) {
        var d = app.activeDocument;
        var l = d.artLayers.add();
        unclip(l);
        l.kind = LayerKind.TEXT;
        l.name = name;
        var t = l.textItem;
        t.kind = TextType.PARAGRAPHTEXT;
        t.contents = contents;
        t.font = F[o.f];
        t.size = new UnitValue(o.s, "px");
        t.color = rgb(o.c);
        if (o.ld) {
            t.useAutoLeading = false;
            t.leading = new UnitValue(o.ld, "px");
        }
        if (o.tr) { t.tracking = o.tr; }
        if (o.caps) { t.capitalization = TextCase.ALLCAPS; }
        if (o.j === "R") { t.justification = Justification.RIGHT; }
        if (o.j === "C") { t.justification = Justification.CENTER; }
        t.position = [new UnitValue(x, "px"), new UnitValue(y, "px")];
        t.width = new UnitValue(w, "px");
        t.height = new UnitValue(h, "px");
        if (o.hidden) { l.visible = false; }
        return l;
    }

    // ---- shared pieces ----
    function eyebrow(name, str, x, y, w, color) {
        return text(name, str, x, y, w, 40, { f: "SANSB", s: 28, ld: 34, c: color, tr: 143, caps: true });
    }
    function label(name, str, x, y, w, color) {
        return text(name, str, x, y, w, 36, { f: "SANSB", s: 24, ld: 30, c: color, tr: 150, caps: true });
    }
    function heading(name, str, x, y, w, color) {
        return text(name, str, x, y, w, 90, { f: "OSW", s: 72, ld: 80, c: color, caps: true });
    }
    function footer(num, color, leftText) {
        if (leftText !== null) {
            text("Footer", leftText || "LL Johnson \u00D7 Tyler Kosec Media", 128, 986, 1100, 36, { f: "SANS", s: 24, ld: 30, c: color });
        }
        text("Page number", num, 1592, 986, 200, 36, { f: "SANS", s: 24, ld: 30, c: color, j: "R" });
    }

    function newSlide(name) {
        var d = app.documents.add(W, H, 72, name, NewDocumentMode.RGB, DocumentFill.WHITE);
        return d;
    }

    function saveSlide(doc, fileBase) {
        var f = new File(outFolder.fsName + "/" + fileBase + ".psd");
        var o = new PhotoshopSaveOptions();
        o.layers = true;
        o.embedColorProfile = true;
        doc.saveAs(f, o, false, Extension.LOWERCASE);
        savedFiles.push(f);
        try {
            var j = new JPEGSaveOptions();
            j.quality = 11;
            j.embedColorProfile = true;
            doc.saveAs(new File(outFolder.fsName + "/" + fileBase + ".jpg"), j, true, Extension.LOWERCASE);
        } catch (e) {}
    }

    // ---- slides ----
    function slideCover() {
        var d = newSlide("01 Cover");
        bg(DARK);
        photo("Cover", 0, 0, 1920, 1080);
        scrim("Scrim (gradient, double-click to edit)", DARK, [[0, 95], [36, 86], [60, 30], [76, 0], [100, 0]]);
        rect("Accent bar", 0, 0, 24, 1080, ACCENT);
        eyebrow("Eyebrow", "A content plan for LL Johnson Distributing", 128, 233, 1500, ACCENT);
        text("Title", "Who makes\rit run.", 128, 307, 1100, 380, { f: "OSW", s: 180, ld: 180, c: LIGHT, caps: true });
        text("Subtitle", "A month-by-month plan for LL Johnson's channels. I shoot it and post it. You keep the final say.",
            128, 707, 920, 180, { f: "SANS", s: 40, ld: 54, c: BODY_D });
        text("Footer", "Prepared for John Knott Jr.  \u00B7  Tyler Kosec Media  \u00B7  October 2026", 128, 986, 1100, 36,
            { f: "SANS", s: 24, ld: 30, c: BODY_D });
        saveSlide(d, "01-cover");
    }

    function slideInsight() {
        var d = newSlide("02 Insight");
        bg(LIGHT);
        photo("Right photo", 1180, 0, 740, 1080);
        rect("OPTIONAL caption box (turn on with the caption)", 1212, 946, 636, 70, DARK, 82, true);
        text("OPTIONAL caption (edit, then turn on)", "Caption: where and when this was shot", 1232, 962, 600, 40,
            { f: "SANS", s: 24, ld: 30, c: LIGHT, hidden: true });
        eyebrow("Eyebrow", "What I noticed", 128, 147, 960, ACCENT_DK);
        text("Headline 1", "The superintendent gets the credit.", 128, 217, 960, 175, { f: "OSW", s: 76, ld: 84, c: DARK, caps: true });
        text("Headline 2", "The equipment manager makes it run.", 128, 385, 960, 175, { f: "OSW", s: 76, ld: 84, c: ACCENT_DK, caps: true });
        text("Body", "Every stripe on a field is somebody's math. Across Colorado, when a machine goes down the week of a tournament, that's the person who calls LL Johnson.",
            128, 589, 960, 210, { f: "SANS", s: 34, ld: 48, c: BODY_L });
        text("Closing line", "I've seen firsthand how LL Johnson shows up for that person. Most of it never gets filmed.",
            128, 817, 960, 100, { f: "SANSSB", s: 30, ld: 42, c: DARK });
        text("Footer", "LL Johnson \u00D7 Tyler Kosec Media", 128, 986, 760, 36, { f: "SANS", s: 24, ld: 30, c: BODY_L });
        text("Page number", "02", 940, 986, 148, 36, { f: "SANS", s: 24, ld: 30, c: BODY_L, j: "R" });
        saveSlide(d, "02-insight");
    }

    function statRow(prefix, num, lab, x, y, numColor, labColor, labH, hidden) {
        text(prefix + " number", num, x, y, 200, 76, { f: "OSW", s: 60, ld: 66, c: numColor, hidden: hidden });
        text(prefix + " label", lab, x + 220, y, 504, labH || 84, { f: "SANS", s: 28, ld: 38, c: labColor, hidden: hidden });
    }

    function slideGap() {
        var d = newSlide("03 TikTok");
        bg(LIGHT2);
        heading("Headline 1", "People already watch your machines.", 128, 128, 1664, DARK);
        heading("Headline 2", "Now show who runs them.", 128, 208, 1664, ACCENT_DK);
        rect("Left card", 128, 332, 812, 608, DARK);
        rect("Right card", 980, 332, 812, 608, LIGHT);
        label("Left card label", "What's working: TikTok", 172, 376, 724, ACCENT);
        statRow("Row 1", "2,000+", "likes on the Toro Dingo TXL 2000 at ProGreen Expo (Feb 2026)", 172, 440, ACCENT, LIGHT);
        statRow("Row 2", "1,200+", "likes on the GeoLink autonomous fairway mower (Jun 2025)", 172, 544, ACCENT, LIGHT);
        statRow("OPTIONAL Row 3 (fill in, then turn on)", "[#]", "TikTok followers", 172, 648, ACCENT, LIGHT, 84, true);
        text("Left closing line", "Machines in action do best so far. The audience is there.", 172, 760, 724, 90,
            { f: "SANS", s: 28, ld: 39, c: BODY_D });
        label("Right card label", "Where it can go next", 1024, 376, 724, ACCENT_DK);
        statRow("Row 4", "<100", "LinkedIn followers. The page still carries the Pattlen name from before the sale.", 1024, 440, DARK, BODY_L, 124);
        statRow("Row 5", "2+", "Facebook pages under the LL Johnson name", 1024, 578, DARK, BODY_L);
        statRow("Row 6", "50", "years in 2026, and 22+ Toro awards. A story worth telling bigger.", 1024, 682, DARK, BODY_L, 124);
        text("Right closing line", "Same audience, more channels. It just needs more hands.", 1024, 826, 724, 90,
            { f: "SANS", s: 28, ld: 39, c: BODY_L });
        footer("03", BODY_L, "Public counts as found Oct 9, 2026; live counts may be higher. Award count as of 2018.");
        saveSlide(d, "03-tiktok");
    }

    function slideMission() {
        var d = newSlide("04 Mission");
        bg(ACCENT);
        eyebrow("Eyebrow", "The mission", 128, 158, 1664, DARK);
        text("Mission", "Put the equipment manager in the spotlight.\rShow that LL\u00A0Johnson is in their corner.",
            128, 248, 1664, 490, { f: "OSW", s: 110, ld: 116, c: DARK, caps: true });
        rect("Rule left", 128, 768, 800, 4, DARK);
        rect("Rule right", 992, 768, 800, 4, DARK);
        text("Label left", "For equipment managers", 128, 796, 800, 36, { f: "SANSB", s: 26, ld: 32, c: BROWN, tr: 115, caps: true });
        text("Label right", "For everyone else", 992, 796, 800, 36, { f: "SANSB", s: 26, ld: 32, c: BROWN, tr: 115, caps: true });
        text("Statement left", "You're seen. You're backed.", 128, 840, 800, 60, { f: "SANSSB", s: 40, ld: 52, c: DARK });
        text("Statement right", "This is who makes it run.", 992, 840, 800, 60, { f: "SANSSB", s: 40, ld: 52, c: DARK });
        footer("04", DARK, null);
        saveSlide(d, "04-mission");
    }

    function slidePillars() {
        var d = newSlide("05 Pillars");
        bg(LIGHT);
        heading("Headline", "Four kinds of posts. One story.", 128, 128, 1664, DARK);
        var P = [
            ["01", "The Bench", "The craft: grinding, reel setup, Dingo and Ventrac checks, winterizing.", "About 4 a month"],
            ["02", "Spotlights", "Equipment managers and techs at Colorado courses, fields and parks.", "About 4 a month"],
            ["03", "In Your Corner", "LL Johnson showing up: parts runs, techs on site, education.", "2\u20133 a month"],
            ["04", "This Week in the Shop", "What Colorado crews are doing now: irrigation blowouts, snow and ice prep, new Toro machines.", "1\u20132 a month"]
        ];
        for (var i = 0; i < P.length; i++) {
            var x = 128 + i * 424;
            var n = "Pillar " + P[i][0];
            rect(n + " card", x, 256, 392, 470, LIGHT2);
            rect(n + " top bar", x, 256, 392, 8, ACCENT);
            text(n + " number", P[i][0], x + 32, 296, 328, 32, { f: "SANSB", s: 24, ld: 30, c: ACCENT_DK });
            text(n + " title", P[i][1], x + 32, 336, 328, 104, { f: "OSW", s: 44, ld: 48, c: DARK, caps: true });
            text(n + " text", P[i][2], x + 32, 446, 328, 210, { f: "SANS", s: 28, ld: 38, c: BODY_L });
            text(n + " count", P[i][3], x + 32, 676, 328, 32, { f: "SANSB", s: 24, ld: 30, c: DARK });
        }
        label("Test label", "Every post passes one test", 128, 774, 1664, ACCENT_DK);
        text("Test question", "Could only LL Johnson post this?", 128, 812, 1664, 60, { f: "SANSSB", s: 40, ld: 52, c: DARK });
        footer("05", BODY_L);
        saveSlide(d, "05-pillars");
    }

    function slideSeries() {
        var d = newSlide("06 Series");
        bg(DARK);
        heading("Headline", "Four series that carry it.", 128, 128, 1664, LIGHT);
        var S = [
            ["Spotlights", "My EM", "A superintendent gives the credit away: what their equipment manager does that nobody sees."],
            ["The Bench", "The Math Behind the Cut", "Why the greens roll fast, one number at a time. Made for golfers and fans, so the EMs and supers behind it share it."],
            ["In Your Corner", "Saved the Day", "Customers tell the story of LL Johnson coming through. Their words, not ours."],
            ["Spotlights", "EM of the Month", "Colorado nominates its own. LL Johnson backs the EM community that's already forming."]
        ];
        for (var i = 0; i < S.length; i++) {
            var x = (i % 2 === 0) ? 128 : 976;
            var y = (i < 2) ? 256 : 556;
            var n = "Series " + (i + 1);
            rect(n + " card", x, y, 816, 270, DARK2);
            label(n + " pillar tag", S[i][0], x + 40, y + 36, 736, BODY_D);
            text(n + " title", S[i][1], x + 40, y + 74, 736, 62, { f: "OSW", s: 48, ld: 53, c: ACCENT, caps: true });
            text(n + " text", S[i][2], x + 40, y + 138, 736, 124, { f: "SANS", s: 28, ld: 39, c: BODY_D });
        }
        footer("06", BODY_D);
        saveSlide(d, "06-series");
    }

    function slideRoom() {
        var d = newSlide("07 Proof");
        bg(DARK);
        heading("Headline", "Already in the room.", 128, 128, 1664, LIGHT);
        photo("Big photo", 128, 244, 1080, 510);
        photo("Top right photo", 1232, 244, 560, 243);
        photo("Bottom right photo", 1232, 511, 560, 243);
        text("Caption", "Stills from my Toro shoots. This is the eye I'd bring to LL Johnson every month.",
            128, 790, 1200, 96, { f: "SANS", s: 30, ld: 42, c: BODY_D });
        rect("Button", 1420, 800, 372, 68, ACCENT);
        text("Button text", "See my work \u2192", 1420, 818, 372, 40, { f: "SANSB", s: 30, ld: 36, c: DARK, j: "C" });
        footer("07", BODY_D);
        saveSlide(d, "07-proof");
    }

    function slideMonth() {
        var d = newSlide("08 Month");
        bg(LIGHT2);
        heading("Headline", "You keep the voice. I add the hands.", 128, 128, 1664, DARK);
        var M = [
            ["12", "finished posts a month, each cut to fit Instagram, Facebook, LinkedIn and TikTok"],
            ["2", "shoot days a month at branches, courses and fields"],
            ["1", "monthly report: what worked and what's next"]
        ];
        for (var i = 0; i < M.length; i++) {
            var x = 128 + i * 565;
            var n = "Stat " + (i + 1);
            rect(n + " rule", x, 256, 533, 4, DARK);
            text(n + " number", M[i][0], x, 276, 533, 112, { f: "OSW", s: 96, ld: 106, c: ACCENT_DK });
            text(n + " label", M[i][1], x, 392, 533, 124, { f: "SANS", s: 28, ld: 38, c: BODY_L });
        }
        rect("I handle box", 128, 546, 816, 200, LIGHT);
        rect("You keep box", 976, 546, 816, 200, DARK);
        label("I handle label", "I handle", 168, 586, 736, ACCENT_DK);
        label("You keep label", "You keep", 1016, 586, 736, ACCENT);
        text("I handle text", "Planning, shooting, editing, captions, posting and reporting.", 168, 628, 736, 96,
            { f: "SANS", s: 30, ld: 42, c: DARK });
        text("You keep text", "Final say on every post, the customer relationships, and the voice.", 1016, 628, 736, 96,
            { f: "SANS", s: 30, ld: 42, c: LIGHT });
        text("From you", "From you: about an hour a week. Approve the month's plan, open a door or two, and give a yes or no on each post.",
            128, 782, 1664, 96, { f: "SANSSB", s: 30, ld: 42, c: DARK });
        footer("08", BODY_L);
        saveSlide(d, "08-month");
    }

    function slideFirst90() {
        var d = newSlide("09 First 90 days");
        bg(LIGHT);
        heading("Headline", "The first 90 days.", 128, 128, 1664, DARK);
        var C = [
            ["Days 1\u201330", "Get set up together", "Agree the first month's plan. Take a baseline of every channel. Talk through the LinkedIn name and extra Facebook pages with whoever owns that call. First shoot day."],
            ["Days 31\u201360", "Launch the series", "The Bench goes weekly. My EM every other week. Open the first EM of the Month nominations."],
            ["Days 61\u201390", "Read it, double down", "Compare against day one. Keep what works. Build the spring startup calendar."]
        ];
        for (var i = 0; i < C.length; i++) {
            var x = 128 + i * 565;
            var n = "Phase " + (i + 1);
            rect(n + " top bar", x, 252, 533, 8, ACCENT);
            rect(n + " card", x, 260, 533, 460, LIGHT2);
            label(n + " label", C[i][0], x + 36, 296, 461, ACCENT_DK);
            text(n + " title", C[i][1], x + 36, 336, 461, 104, { f: "OSW", s: 44, ld: 48, c: DARK, caps: true });
            text(n + " text", C[i][2], x + 36, 444, 461, 250, { f: "SANS", s: 28, ld: 39, c: BODY_L });
        }
        rect("Measure rule", 128, 752, 6, 124, DARK);
        label("Measure label", "How we'll know it's working", 162, 756, 1630, ACCENT_DK);
        text("Measure text", "Shares and tags from EMs and superintendents  \u00B7  saves on how-to posts  \u00B7  EM of the Month nominations  \u00B7  a simple tally at the counter: \"saw it online.\"",
            162, 794, 1630, 96, { f: "SANS", s: 30, ld: 42, c: DARK });
        footer("09", BODY_L);
        saveSlide(d, "09-first-90-days");
    }

    function slideNext() {
        var d = newSlide("10 Next step");
        bg(DARK);
        photo("Background photo", 0, 0, 1920, 1080);
        scrim("Scrim (gradient, double-click to edit)", DARK, [[0, 95], [50, 88], [100, 55]]);
        rect("Accent bar", 0, 0, 24, 1080, ACCENT);
        eyebrow("Eyebrow", "Next step", 128, 175, 1500, ACCENT);
        text("Headline", "Let's pick the\rfirst three shoots.", 128, 249, 1600, 300, { f: "OSW", s: 140, ld: 143, c: LIGHT, caps: true });
        text("Subhead", "One 20-minute call. You bring the people and the stories. I'll bring the plan and the monthly cost, so you have something to share internally.",
            128, 575, 1300, 175, { f: "SANS", s: 40, ld: 54, c: BODY_D });
        text("Reply line", "Reply to this email with a day next week that works.", 128, 765, 1300, 50, { f: "SANSSB", s: 32, ld: 40, c: LIGHT });
        rect("Rule", 128, 833, 1300, 2, LINE_D);
        text("Name", "Tyler Kosec  \u00B7  Tyler Kosec Media  \u00B7  Loveland, CO", 128, 859, 1300, 44, { f: "SANSB", s: 32, ld: 40, c: LIGHT });
        text("Contact (add your phone)", "tylerkosecmedia@gmail.com  \u00B7  [phone]", 128, 905, 1300, 42, { f: "SANS", s: 30, ld: 38, c: BODY_D });
        saveSlide(d, "10-next-step");
    }

    // ---- run ----
    var stillsFolder = Folder.selectDialog("Step 1 of 2: pick the folder with your stills (F:\\TORO\\STILLS). Cancel to leave the photo frames empty.");
    var outFolder = Folder.selectDialog("Step 2 of 2: pick the folder to save the 10 slide PSDs into.");
    if (!outFolder) {
        alert("No save folder picked, so nothing was built.");
        return;
    }
    if (stillsFolder) {
        var all = stillsFolder.getFiles();
        for (var i = 0; i < all.length; i++) {
            if (all[i] instanceof File && /\.(jpe?g|png|tiff?|psd|webp)$/i.test(all[i].name)) { stills.push(all[i]); }
        }
        stills.sort(function (a, b) {
            var x = a.name.toLowerCase(), y = b.name.toLowerCase();
            return x < y ? -1 : (x > y ? 1 : 0);
        });
    }

    try {
        app.preferences.rulerUnits = Units.PIXELS;
        app.preferences.typeUnits = TypeUnits.PIXELS;
        app.displayDialogs = DialogModes.NO;
        loadFonts();

        slideCover();
        slideInsight();
        slideGap();
        slideMission();
        slidePillars();
        slideSeries();
        slideRoom();
        slideMonth();
        slideFirst90();
        slideNext();

        var pdfNote = "";
        try {
            var po = new PresentationOptions();
            po.presentation = false;
            po.view = false;
            po.includeFilename = false;
            var pdfOpts = new PDFSaveOptions();
            pdfOpts.jpegQuality = 10;
            pdfOpts.layers = false;
            po.PDFFileOptions = pdfOpts;
            app.makePDFPresentation(savedFiles, new File(outFolder.fsName + "/LL Johnson Pitch - Tyler Kosec Media.pdf"), po);
            pdfNote = "\nAlso made: LL Johnson Pitch - Tyler Kosec Media.pdf";
        } catch (e2) {
            pdfNote = "\n(The PDF step didn't run. Use File > Automate > PDF Presentation on the 10 PSDs.)";
        }

        var msg = "Done. 10 slides saved as PSD + JPG in:\n" + outFolder.fsName + pdfNote +
            "\n\nStills placed: " + placedCount + " of " + frameCount + " photo frames.";
        if (stills.length > frameCount) { msg += "\n(" + (stills.length - frameCount) + " more stills in the folder were not used.)"; }
        var miss = [];
        for (var m in MISSING) { if (MISSING.hasOwnProperty(m)) { miss.push(m); } }
        if (miss.length) {
            msg += "\n\nFonts not found, Arial used instead: " + miss.join(", ") +
                ".\nInstall Oswald and Public Sans (free at fonts.google.com), then rerun the script.";
        }
        alert(msg);
    } catch (err) {
        alert("Something went wrong building the slides:\n" + err + (err.line ? "\nLine: " + err.line : ""));
    } finally {
        restore();
    }
})();
