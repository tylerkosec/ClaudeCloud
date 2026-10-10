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
        eyebrow("Eyebrow", "LL Johnson \u00D7 Tyler Kosec Media", 128, 260, 1500, ACCENT);
        text("Title", "Who makes\rit run.", 128, 334, 1100, 380, { f: "OSW", s: 180, ld: 180, c: LIGHT, caps: true });
        text("Subtitle", "A content plan for LL Johnson. I shoot it, edit it and post it. You approve it.",
            128, 734, 920, 120, { f: "SANS", s: 40, ld: 54, c: BODY_D });
        text("Footer", "For John Knott Jr.  \u00B7  October 2026", 128, 986, 1100, 36,
            { f: "SANS", s: 24, ld: 30, c: BODY_D });
        saveSlide(d, "01-cover");
    }

    function slideInsight() {
        var d = newSlide("02 Insight");
        bg(LIGHT);
        photo("Right photo", 1180, 0, 740, 1080);
        rect("OPTIONAL caption box (turn on with the caption)", 1212, 946, 636, 70, DARK, 82, true);
        text("OPTIONAL caption (edit, then turn on)", "Where and when this was shot", 1232, 962, 600, 40,
            { f: "SANS", s: 24, ld: 30, c: LIGHT, hidden: true });
        text("Headline 1", "Supers get the credit.", 128, 248, 960, 90, { f: "OSW", s: 76, ld: 84, c: DARK, caps: true });
        text("Headline 2", "Equipment managers keep it running.", 128, 332, 960, 175, { f: "OSW", s: 76, ld: 84, c: ACCENT_DK, caps: true });
        text("Body", "Every stripe on a field is somebody's math. When a mower goes down before a tournament, the EM is the one calling LL Johnson.",
            128, 536, 960, 160, { f: "SANS", s: 34, ld: 48, c: BODY_L });
        text("Closing line", "I've seen how you show up for them. Most of it never gets filmed.",
            128, 716, 960, 100, { f: "SANSSB", s: 30, ld: 42, c: DARK });
        text("Footer", "LL Johnson \u00D7 Tyler Kosec Media", 128, 986, 760, 36, { f: "SANS", s: 24, ld: 30, c: BODY_L });
        text("Page number", "02", 940, 986, 148, 36, { f: "SANS", s: 24, ld: 30, c: BODY_L, j: "R" });
        saveSlide(d, "02-insight");
    }

    function statRow(prefix, num, lab, x, y, numColor, labColor, labH, hidden) {
        text(prefix + " number", num, x, y, 200, 76, { f: "OSW", s: 60, ld: 66, c: numColor, hidden: hidden });
        text(prefix + " label", lab, x + 220, y + 4, 504, labH || 84, { f: "SANS", s: 28, ld: 38, c: labColor, hidden: hidden });
    }

    function slideGap() {
        var d = newSlide("03 TikTok");
        bg(LIGHT2);
        heading("Headline", "People already watch your machines.", 128, 128, 1664, DARK);
        rect("Left card", 128, 256, 812, 440, DARK);
        rect("Right card", 980, 256, 812, 440, LIGHT);
        label("Left card label", "TikTok", 172, 300, 724, ACCENT);
        statRow("Row 1", "2,000+", "likes: Dingo TXL 2000 at ProGreen (Feb 2026)", 172, 360, ACCENT, LIGHT);
        statRow("Row 2", "1,200+", "likes: GeoLink autonomous fairway mower (June 2025)", 172, 464, ACCENT, LIGHT);
        statRow("OPTIONAL Row 3 (fill in, then turn on)", "[#]", "TikTok followers", 172, 568, ACCENT, LIGHT, 84, true);
        label("Right card label", "Room to grow", 1024, 300, 724, ACCENT_DK);
        statRow("Row 4", "<100", "LinkedIn followers, still under the Pattlen name", 1024, 360, DARK, BODY_L);
        statRow("Row 5", "2+", "Facebook pages for LL Johnson", 1024, 464, DARK, BODY_L);
        statRow("Row 6", "50", "years in business this year", 1024, 568, DARK, BODY_L);
        text("Next line", "Next: put the people running those machines on camera.", 128, 744, 1664, 60,
            { f: "SANSSB", s: 40, ld: 52, c: ACCENT_DK });
        footer("03", BODY_L, "Public counts, Oct 9, 2026.");
        saveSlide(d, "03-tiktok");
    }

    function slideMission() {
        var d = newSlide("04 Mission");
        bg(ACCENT);
        eyebrow("Eyebrow", "The mission", 128, 227, 1664, DARK);
        text("Mission", "Make the equipment manager the hero.\rShow LL\u00A0Johnson has their back.",
            128, 317, 1664, 520, { f: "OSW", s: 120, ld: 126, c: DARK, caps: true });
        footer("04", DARK, null);
        saveSlide(d, "04-mission");
    }

    function slidePillars() {
        var d = newSlide("05 Pillars");
        bg(LIGHT);
        heading("Headline", "What we post", 128, 128, 1664, DARK);
        var P = [
            ["The Bench", "Shop work. Grinding, reel setup, Dingo and Ventrac service.", "About 4 a month"],
            ["Spotlights", "EMs and techs at Colorado courses, fields and parks.", "About 4 a month"],
            ["In Your Corner", "LL Johnson coming through. Parts runs, techs on site, training.", "2\u20133 a month"],
            ["This Week", "What crews are doing now. Blowouts, snow prep, new Toro gear.", "1\u20132 a month"]
        ];
        for (var i = 0; i < P.length; i++) {
            var x = 128 + i * 424;
            var n = "Pillar " + (i + 1);
            rect(n + " card", x, 256, 392, 420, LIGHT2);
            rect(n + " top bar", x, 256, 392, 8, ACCENT);
            text(n + " title", P[i][0], x + 32, 300, 328, 60, { f: "OSW", s: 44, ld: 48, c: DARK, caps: true });
            text(n + " text", P[i][1], x + 32, 372, 328, 200, { f: "SANS", s: 28, ld: 38, c: BODY_L });
            text(n + " count", P[i][2], x + 32, 620, 328, 32, { f: "SANSB", s: 24, ld: 30, c: DARK });
        }
        label("Rule label", "The rule", 128, 732, 1664, ACCENT_DK);
        text("Rule", "If another dealer could post it, we don't.", 128, 770, 1664, 60, { f: "SANSSB", s: 40, ld: 52, c: DARK });
        footer("05", BODY_L);
        saveSlide(d, "05-pillars");
    }

    function slideSeries() {
        var d = newSlide("06 Series");
        bg(DARK);
        heading("Headline", "Four series", 128, 128, 1664, LIGHT);
        var S = [
            ["My EM", "A super tells you what their EM does that nobody sees."],
            ["The Math Behind the Cut", "Why the greens roll fast, one number at a time."],
            ["Saved the Day", "Customers tell how LL Johnson came through."],
            ["EM of the Month", "Colorado nominates. LL Johnson backs it."]
        ];
        for (var i = 0; i < S.length; i++) {
            var x = (i % 2 === 0) ? 128 : 976;
            var y = (i < 2) ? 256 : 540;
            var n = "Series " + (i + 1);
            rect(n + " card", x, y, 816, 252, DARK2);
            text(n + " title", S[i][0], x + 40, y + 56, 736, 62, { f: "OSW", s: 48, ld: 53, c: ACCENT, caps: true });
            text(n + " text", S[i][1], x + 40, y + 128, 736, 90, { f: "SANS", s: 30, ld: 42, c: BODY_D });
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
        text("Caption", "Stills from my Toro shoots.", 128, 812, 1200, 50, { f: "SANS", s: 30, ld: 42, c: BODY_D });
        rect("Button", 1420, 800, 372, 68, ACCENT);
        text("Button text", "Watch the reel \u2192", 1420, 818, 372, 40, { f: "SANSB", s: 30, ld: 36, c: DARK, j: "C" });
        footer("07", BODY_D);
        saveSlide(d, "07-proof");
    }

    function slideMonth() {
        var d = newSlide("08 Month");
        bg(LIGHT2);
        heading("Headline", "You keep the voice. I do the work.", 128, 128, 1664, DARK);
        var M = [
            ["12", "posts a month, cut for Instagram, Facebook, LinkedIn and TikTok"],
            ["2", "shoot days a month"],
            ["1", "report a month: what worked, what's next"]
        ];
        for (var i = 0; i < M.length; i++) {
            var x = 128 + i * 565;
            var n = "Stat " + (i + 1);
            rect(n + " rule", x, 256, 533, 4, DARK);
            text(n + " number", M[i][0], x, 276, 533, 112, { f: "OSW", s: 96, ld: 106, c: ACCENT_DK });
            text(n + " label", M[i][1], x, 392, 533, 124, { f: "SANS", s: 28, ld: 38, c: BODY_L });
        }
        rect("I handle box", 128, 556, 816, 210, LIGHT);
        rect("You box", 976, 556, 816, 210, DARK);
        label("I handle label", "I handle", 168, 596, 736, ACCENT_DK);
        label("You label", "You", 1016, 596, 736, ACCENT);
        text("I handle text", "Planning, shooting, editing, captions, posting.", 168, 638, 736, 96,
            { f: "SANS", s: 30, ld: 42, c: DARK });
        text("You text", "Approve the plan, open a few doors, OK each post. About an hour a week.", 1016, 638, 736, 96,
            { f: "SANS", s: 30, ld: 42, c: LIGHT });
        footer("08", BODY_L);
        saveSlide(d, "08-month");
    }

    function slideFirst90() {
        var d = newSlide("09 First 90 days");
        bg(LIGHT);
        heading("Headline", "The first 90 days", 128, 128, 1664, DARK);
        var C = [
            ["Month 1", "Set up", "Plan the month, baseline every channel, first shoot. Sort out the LinkedIn name and extra Facebook pages with whoever owns them."],
            ["Month 2", "Launch", "The Bench every week. My EM every other week. First EM of the Month."],
            ["Month 3", "Double down", "Compare to month one. Keep what works. Plan spring startup."]
        ];
        for (var i = 0; i < C.length; i++) {
            var x = 128 + i * 565;
            var n = "Month " + (i + 1);
            rect(n + " top bar", x, 252, 533, 8, ACCENT);
            rect(n + " card", x, 260, 533, 420, LIGHT2);
            label(n + " label", C[i][0], x + 36, 296, 461, ACCENT_DK);
            text(n + " title", C[i][1], x + 36, 336, 461, 60, { f: "OSW", s: 44, ld: 48, c: DARK, caps: true });
            text(n + " text", C[i][2], x + 36, 404, 461, 250, { f: "SANS", s: 28, ld: 39, c: BODY_L });
        }
        rect("Measure rule", 128, 716, 6, 124, DARK);
        label("Measure label", "How we'll know", 162, 720, 1630, ACCENT_DK);
        text("Measure text", "EMs and supers sharing and tagging, saves on how-to posts, nominations, and the counter hearing \"saw it online.\"",
            162, 758, 1630, 96, { f: "SANS", s: 30, ld: 42, c: DARK });
        footer("09", BODY_L);
        saveSlide(d, "09-first-90-days");
    }

    function slideNext() {
        var d = newSlide("10 Next step");
        bg(DARK);
        photo("Background photo", 0, 0, 1920, 1080);
        scrim("Scrim (gradient, double-click to edit)", DARK, [[0, 95], [50, 88], [100, 55]]);
        rect("Accent bar", 0, 0, 24, 1080, ACCENT);
        text("Headline", "Let's pick the\rfirst three shoots.", 128, 227, 1600, 300, { f: "OSW", s: 140, ld: 143, c: LIGHT, caps: true });
        text("Subhead", "20-minute call. I'll bring the plan and the monthly price.",
            128, 553, 1300, 60, { f: "SANS", s: 40, ld: 54, c: BODY_D });
        text("Reply line", "Reply with a day next week that works.", 128, 631, 1300, 50, { f: "SANSSB", s: 32, ld: 40, c: LIGHT });
        rect("Rule", 128, 711, 1300, 2, LINE_D);
        text("Name", "Tyler Kosec  \u00B7  Tyler Kosec Media  \u00B7  Loveland, CO", 128, 737, 1300, 44, { f: "SANSB", s: 32, ld: 40, c: LIGHT });
        text("Contact (add your phone)", "tylerkosecmedia@gmail.com  \u00B7  [phone]", 128, 783, 1300, 42, { f: "SANS", s: 30, ld: 38, c: BODY_D });
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
