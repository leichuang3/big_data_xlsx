var strs = {}; // shared strings
var _ssfopts = {}; // spreadsheet formatting options


/*global Map */
var browser_has_Map = typeof Map !== 'undefined';

function get_sst_id(sst/*:SST*/, str/*:string*/, rev)/*:number*/ {
	var i = 0, len = sst.length;
	if(rev) {
		if(browser_has_Map ? rev.has(str) : Object.prototype.hasOwnProperty.call(rev, str)) {
			var revarr = browser_has_Map ? rev.get(str) : rev[str];
			for(; i < revarr.length; ++i) {
				if(sst[revarr[i]].t === str) { sst.Count ++; return revarr[i]; }
			}
		}
	} else for(; i < len; ++i) {
		if(sst[i].t === str) { sst.Count ++; return i; }
	}
	sst[len] = ({t:str}/*:any*/); sst.Count ++; sst.Unique ++;
	if(rev) {
		if(browser_has_Map) {
			if(!rev.has(str)) rev.set(str, []);
			rev.get(str).push(len);
		} else {
			if(!Object.prototype.hasOwnProperty.call(rev, str)) rev[str] = [];
			rev[str].push(len);
		}
	}
	return len;
}

function col_obj_w(C/*:number*/, col) {
	var p = ({min:C+1,max:C+1}/*:any*/);
	/* wch (chars), wpx (pixels) */
	var wch = -1;
	if(col.MDW) MDW = col.MDW;
	if(col.width != null) p.customWidth = 1;
	else if(col.wpx != null) wch = px2char(col.wpx);
	else if(col.wch != null) wch = col.wch;
	if(wch > -1) { p.width = char2width(wch); p.customWidth = 1; }
	else if(col.width != null) p.width = col.width;
	if(col.hidden) p.hidden = true;
	if(col.level != null) { p.outlineLevel = p.level = col.level; }
	return p;
}

function default_margins(margins/*:Margins*/, mode/*:?string*/) {
	if(!margins) return;
	var defs = [0.7, 0.7, 0.75, 0.75, 0.3, 0.3];
	if(mode == 'xlml') defs = [1, 1, 1, 1, 0.5, 0.5];
	if(margins.left   == null) margins.left   = defs[0];
	if(margins.right  == null) margins.right  = defs[1];
	if(margins.top    == null) margins.top    = defs[2];
	if(margins.bottom == null) margins.bottom = defs[3];
	if(margins.header == null) margins.header = defs[4];
	if(margins.footer == null) margins.footer = defs[5];
}

function get_cell_style(styles/*:Array<any>*/, cell/*:Cell*/, opts) {
	var z = opts.revssf[cell.z != null ? cell.z : "General"];
	var i = 0x3c, len = styles.length;
	if(z == null && opts.ssf) {
		for(; i < 0x188; ++i) if(opts.ssf[i] == null) {
			SSF__load(cell.z, i);
			// $FlowIgnore
			opts.ssf[i] = cell.z;
			opts.revssf[cell.z] = z = i;
			break;
		}
	}
	for(i = 0; i != len; ++i) if(styles[i].numFmtId === z) return i;
	styles[len] = {
		numFmtId:z,
		fontId:0,
		fillId:0,
		borderId:0,
		xfId:0,
		applyNumberFormat:1
	};
	return len;
}

function safe_format(p/*:Cell*/, fmtid/*:number*/, fillid/*:?number*/, opts, themes, styles) {
	try {
		if(opts.cellNF) p.z = table_fmt[fmtid];
	} catch(e) { if(opts.WTF) throw e; }
	if(p.t === 'z' && !opts.cellStyles) return;
	if(p.t === 'd' && typeof p.v === 'string') p.v = parseDate(p.v);
	if((!opts || opts.cellText !== false) && p.t !== 'z') try {
		if(table_fmt[fmtid] == null) SSF__load(SSFImplicit[fmtid] || "General", fmtid);
		if(p.t === 'e') p.w = p.w || BErr[p.v];
		else if(fmtid === 0) {
			if(p.t === 'n') {
				if((p.v|0) === p.v) p.w = p.v.toString(10);
				else p.w = SSF_general_num(p.v);
			}
			else if(p.t === 'd') {
				var dd = datenum(p.v);
				if((dd|0) === dd) p.w = dd.toString(10);
				else p.w = SSF_general_num(dd);
			}
			else if(p.v === undefined) return "";
			else p.w = SSF_general(p.v,_ssfopts);
		}
		else if(p.t === 'd') p.w = SSF_format(fmtid,datenum(p.v),_ssfopts);
		else p.w = SSF_format(fmtid,p.v,_ssfopts);
	} catch(e) { if(opts.WTF) throw e; }
	if(!opts.cellStyles) return;
	if(fillid != null) try {
		p.s = styles.Fills[fillid];
		if (p.s.fgColor && p.s.fgColor.theme && !p.s.fgColor.rgb) {
			p.s.fgColor.rgb = rgb_tint(themes.themeElements.clrScheme[p.s.fgColor.theme].rgb, p.s.fgColor.tint || 0);
			if(opts.WTF) p.s.fgColor.raw_rgb = themes.themeElements.clrScheme[p.s.fgColor.theme].rgb;
		}
		if (p.s.bgColor && p.s.bgColor.theme) {
			p.s.bgColor.rgb = rgb_tint(themes.themeElements.clrScheme[p.s.bgColor.theme].rgb, p.s.bgColor.tint || 0);
			if(opts.WTF) p.s.bgColor.raw_rgb = themes.themeElements.clrScheme[p.s.bgColor.theme].rgb;
		}
	} catch(e) { if(opts.WTF && styles.Fills) throw e; }
}

function check_ws(ws/*:Worksheet*/, sname/*:string*/, i/*:number*/) {
	if(ws && ws['!ref']) {
		var range = safe_decode_range(ws['!ref']);
		if(range.e.c < range.s.c || range.e.r < range.s.r) throw new Error("Bad range (" + i + "): " + ws['!ref']);
	}
}
function XmlReader(){
	const a = '<![CDATA[';
	const o = ']]>';
	const l = null;
	this.mp = -1;
	this.np = 0;
	this.op = 0;
	this.pp = 0;
	this.buffer = '';
	this.elementType = 2;
	this.depth = 0;
	this.rn = 0;
	this._l = 0;
	this.qp = 0;
	this.rp = 0;
	this.tp = 0;
	this.vp = 0;
	this.lp = 0;
	this.wp = 0;
	this.xmlIndex = 0;
	this.xml = '';
	this.xp = !1;
	this.reset = function(){
		var e = this;
		e.mp = -1;
		e.np = 0;
		e.op = 0;
		e.pp = 0;
		e.buffer = '';
		e.elementType = 2;
		e.depth = 0;
		e.rn = 0;
		e._l = 0;
		e.qp = 0;
		e.rp = 0;
		e.tp = 0;
		e.vp = 0;
		e.lp = 0;
		e.wp = 0;
		e.xmlIndex = 0;
		e.xml = '';
		e.xp = !1;
		e.yp = 0;
		e.zp = '';
	}
	this.setXml = function (e) {
		this.xml = e
	  }
	this.name = function () {
		var e,
		  t = this,
		  r = t.buffer.slice(t.rn, t.rn + t._l)
		return r && !t.keepRootNamespace
		  ? ((e = r.lastIndexOf(':') + 1),
			(e === t.yp && r.substr(0, e) === t.zp) || (e = 0),
			r.substr(e))
		  : r
	  }
	  this.nodeType = function () {
		return 2 === this.elementType ? 15 : 1
	  }
	  this.fillBuffer = function () {
		var e = this,
		  t = e.buffer.length
		return (
		  0 === t && ((e.buffer = e.xml), (e.mp = 0), (e.lp = 0), (e.op = e.buffer.length), !0)
		)
	  }
	  this.read = function () {
		var e,
		  t,
		  r,
		  i,
		  n,
		  l,
		  s,
		  c,
		  u,
		  d,
		  f,
		  p = this
		for (p.np = Number.MAX_VALUE, p.xp = !1; ; ) {
		  if ((p.mp++, p.mp >= p.op && !p.fillBuffer())) return !1
		  if (((e = p.buffer[p.mp]), '<' === e)) break
		}
		for (
		  t = [
			'elementStarting',
			'elementStart',
			'elementNameEnd',
			'elementEnd',
			'elementContent',
			'elementContentStart',
			'endElementStart'
		  ],
			r = t.length,
			i = 0,
			n = !1;
		  i < r;

		)
		  switch (t[i]) {
			case 'elementStarting':
			  for (n = !1; ; ) {
				if ((p.mp++, (e = p.buffer[p.mp]), '/' === e)) {
				  ;(i = 6), (n = !0)
				  break
				}
				if ('?' === e) {
				  for (p.elementType = 3; ; )
					if ((p.mp++, (e = p.buffer[p.mp]), '>' === e)) return !0
				} else if (' ' !== e && '\r' !== e && '\n' !== e && '\t' !== e) {
				  p.rn = p.mp
				  break
				}
			  }
			  if (n) continue
			case 'elementStart':
			  for (n = !1, 1 === p.elementType && p.depth++; ; ) {
				if ((p.mp++, (e = p.buffer[p.mp]), '>' === e)) {
				  ;(p._l = p.mp - p.rn), (i = 3), (n = !0)
				  break
				}
				if (' ' === e || '\r' === e || '\n' === e || '\t' === e || '/' === e) {
				  ;(p._l = p.mp - p.rn), (p.np = p.mp)
				  break
				}
			  }
			  if (
				(0 === p.depth &&
				  !p.keepRootNamespace &&
				  p._l &&
				  ((l = p.buffer.substr(p.rn, p._l)),
				  (s = l.lastIndexOf(':') + 1),
				  s && ((p.zp = l.substr(0, s)), (p.yp = s))),
				n)
			  )
				continue
			case 'elementNameEnd':
			  for (c = !1; ; )
				if ((p.mp++, (e = p.buffer[p.mp]), '"' === e && (c = !c), !c && '>' === e)) break
			  for (u = p.mp; ; ) {
				if ((u--, (e = p.buffer[u]), '/' === e))
				  return (p.pp = u), (p.elementType = 3), !0
				if (' ' !== e && '\r' !== e && '\n' !== e && '\t' !== e) {
				  ;(p.pp = u), (p.elementType = 1), (i = 4)
				  break
				}
			  }
			  continue
			case 'elementEnd':
			  for (u = p.mp; ; ) {
				if ((u--, (e = p.buffer[u]), '/' === e)) return (p.elementType = 3), !0
				if (' ' !== e && '\r' !== e && '\n' !== e && '\t' !== e) {
				  p.elementType = 1
				  break
				}
			  }
			case 'elementContent':
			  for (d = !1; ; ) {
				if ((p.mp++, (e = p.buffer[p.mp]), p.buffer.substr(p.mp, 9) === a))
				  return (
					(f = p.buffer.indexOf(o, p.mp)),
					(p.lp = p.mp),
					(p.wp = f + o.length),
					(p.mp = p.wp - 1),
					(p.xp = !0),
					!0
				  )
				if ('<' === e) return p.mp--, d && (p.wp = p.mp + 1), !0
				if (((p.lp = p.mp), (d = !0), '\r' !== e && '\n' !== e && '\t' !== e)) break
			  }
			case 'elementContentStart':
			  for (;;)
				if ((p.mp++, (e = p.buffer[p.mp]), '<' === e)) return (p.wp = p.mp), p.mp--, !0
			case 'endElementStart':
			  for (
				(2 !== p.elementType && 3 !== p.elementType) || p.depth--,
				  p.elementType = 2,
				  p.rn = p.mp + 1;
				;

			  )
				if ((p.mp++, (e = p.buffer[p.mp]), '>' === e)) return (p._l = p.mp - p.rn), !0
		  }
	  }
	  this.fastRead = function () {
		var e,
		  t,
		  r,
		  i,
		  n,
		  l,
		  s,
		  c,
		  u,
		  d = this
		for (d.np = Number.MAX_VALUE, d.xp = !1; ; ) {
		  if ((d.mp++, d.mp >= d.op && !d.fillBuffer())) return !1
		  if (((e = d.buffer[d.mp]), '<' === e)) break
		}
		for (
		  t = [
			'elementStarting',
			'elementStart',
			'elementNameEnd',
			'elementEnd',
			'elementContent',
			'elementContentStart',
			'endElementStart'
		  ],
			r = 0,
			i = !1;
		  r < t.length;

		)
		  switch (t[r]) {
			case 'elementStarting':
			  for (i = !1; ; ) {
				if ((d.mp++, (e = d.buffer[d.mp]), '/' === e)) {
				  ;(r = 6), (i = !0)
				  break
				}
				if (' ' !== e && '\r' !== e && '\n' !== e && '\t' !== e) {
				  d.rn = d.mp
				  break
				}
			  }
			  if (i) continue
			case 'elementStart':
			  for (i = !1, 1 === d.elementType && d.depth++; ; ) {
				if ((d.mp++, (e = d.buffer[d.mp]), '>' === e)) {
				  ;(d._l = d.mp - d.rn), (r = 3), (i = !0)
				  break
				}
				if (' ' === e || '\r' === e || '\n' === e || '\t' === e) {
				  ;(d._l = d.mp - d.rn), (d.np = d.mp)
				  break
				}
			  }
			  if (
				(0 === d.depth &&
				  !d.keepRootNamespace &&
				  d._l &&
				  ((n = d.buffer.substr(d.rn, d._l)),
				  (l = n.lastIndexOf(':') + 1),
				  l && ((d.zp = n.substr(0, l)), (d.yp = l))),
				i)
			  )
				continue
			case 'elementNameEnd':
			  for (s = !1; ; )
				if ((d.mp++, (e = d.buffer[d.mp]), '"' === e && (s = !s), !s && '>' === e)) break
			  for (c = d.mp; ; ) {
				if ((c--, (e = d.buffer[c]), '/' === e))
				  return (d.pp = c), (d.elementType = 3), !0
				if (' ' !== e && '\r' !== e && '\n' !== e && '\t' !== e) {
				  ;(d.pp = c), (d.elementType = 1), (r = 4)
				  break
				}
			  }
			  continue
			case 'elementEnd':
			  for (c = d.mp; ; ) {
				if ((c--, (e = d.buffer[c]), '/' === e)) return (d.elementType = 3), !0
				if (' ' !== e && '\r' !== e && '\n' !== e && '\t' !== e) {
				  d.elementType = 1
				  break
				}
			  }
			case 'elementContent':
			  for (;;) {
				if ((d.mp++, (e = d.buffer[d.mp]), d.buffer.substr(d.mp, 9) === a))
				  return (
					(u = d.buffer.indexOf(o, d.mp)),
					(d.lp = d.mp),
					(d.wp = u + o.length),
					(d.mp = d.wp - 1),
					(d.xp = !0),
					!0
				  )
				if ('<' === e) return d.mp--, !0
				if ('\r' !== e && '\n' !== e && '\t' !== e) {
				  d.lp = d.mp
				  break
				}
			  }
			case 'elementContentStart':
			  for (;;)
				if ((d.mp++, (e = d.buffer[d.mp]), '<' === e)) return (d.wp = d.mp), d.mp--, !0
			case 'endElementStart':
			  for (
				(2 !== d.elementType && 3 !== d.elementType) || d.depth--, d.elementType = 2;
				;

			  )
				if ((d.mp++, (e = d.buffer[d.mp]), '>' === e)) return !0
		  }
	  }
	  this.moveToNextAttribute = function () {
		for (var e, t, r = this; ; ) {
		  if (r.np >= r.pp - 1) return !1
		  if (
			(r.np++, (e = r.buffer[r.np]), ' ' !== e && '\r' !== e && '\n' !== e && '\t' !== e)
		  ) {
			r.qp = r.np
			break
		  }
		}
		for (;;)
		  if (
			(r.np++,
			(e = r.buffer[r.np]),
			'=' === e || ' ' === e || '\r' === e || '\n' === e || '\t' === e)
		  ) {
			r.rp = r.np - r.qp
			break
		  }
		for (t = '"'; ; )
		  if ((r.np++, (e = r.buffer[r.np]), '"' === e || "'" === e)) {
			;(t = e), (r.tp = r.np + 1)
			break
		  }
		for (;;) if ((r.np++, (e = r.buffer[r.np]), e === t)) return (r.vp = r.np), !0
	  }
	  this.readContentAsString = function () {
		var e = this
		return e.buffer.slice(e.tp, e.vp)
	  }
	  this.readContentAsBoolean = function (e) {
		var t = this,
		  r = t.buffer[t.tp]
		return '1' === r || 't' === r || ('0' !== r && 'f' !== r && e)
	  }
	  this.prototype.readContentAsInt = function (e) {
		var t = this,
		  r = parseInt(t.buffer.slice(t.tp, t.vp), 10)
		return isNaN(r) ? e : r
	  }
	  this.prototype.readContentAsDouble = function (e) {
		var t = this,
		  r = parseFloat(t.buffer.slice(t.tp, t.vp))
		return isNaN(r) ? e : r
	  }
	  this.readContentAsError = function () {
		var e = this.readContentAsString(),
		  t = l
		switch (e) {
		  case '#DIV/0!':
			t = 7
			break
		  case '#N/A':
			t = 42
			break
		  case '#NAME?':
			t = 29
			break
		  case '#NULL!':
			t = 0
			break
		  case '#NUM!':
			t = 36
			break
		  case '#REF!':
			t = 23
			break
		  case '#VALUE!':
			t = 15
			break
		  case '#SPILL!':
			t = 99
		}
		return t !== l ? { _error: e, _code: t } : l
	  }
	  this.readElementContentAsString = function (e) {
		var t,
		  r,
		  i,
		  l = this,
		  s = l.lp
		if (this.lp <= this.rn) return ''
		if (e)
		  for (t = l.buffer, r = t[s - 1]; ' ' === r || '\r' === r || '\n' === r || '\t' === r; )
			s--, (r = t[s - 1])
		return (
		  (i = l.buffer.slice(s, l.wp)),
		  l.xp && ((i = n(i.replace(a, '').replace(o, ''))), (l.xp = !1)),
		  i
		)
	  }
	  this.readElementContentAsInt = function (e) {
		var t = this,
		  r = parseInt(t.buffer.slice(t.lp, t.wp), 10)
		return isNaN(r) ? e : r
	  }
	  this.readElementContentAsDouble = function (e) {
		var t = this,
		  r = parseFloat(t.buffer.slice(t.lp, t.wp))
		return isNaN(r) ? e : r
	  };
	  this.readElementContentAsBoolean = function (e) {
		var t = this,
		  r = t.buffer[t.lp]
		return '1' === r || 't' === r || ('0' !== r && 'f' !== r && e)
	  }
	  this.readAttributeNameAsString = function () {
		var e = this
		return e.buffer.slice(e.qp, e.qp + e.rp)
	  }
	  this.readFullElement = function () {
		var e,
		  t,
		  r = this
		if (2 === r.elementType) return ''
		if (((e = r.rn - 1), 3 === r.elementType)) return r.buffer.slice(e, r.pp + 2)
		if (((t = r.depth), 1 === r.elementType)) {
		  for (; r.read() && !(r.depth <= t); );
		  return r.buffer.slice(e, r.rn + r._l + 1)
		}
		return ''
	  }
}
function a(e, t) {
for (var r, i, n, o = e.depth; e.read() && !(e.depth <= o); )
	if (1 === e.nodeType()) {
	for (
		r = {},
		i = e.name(),
		t[i] ? (Array.isArray(t[i]) || (t[i] = [t[i]]), t[i].push(r)) : (t[i] = r),
		r._attr = {};
		e.moveToNextAttribute();

	)
		r._attr[e.readAttributeNameAsString()] = e.readContentAsString()
	if (3 === e.elementType) continue
	;(n = e.readElementContentAsString()),
		1 === (1 & e.elementType) &&
		'' !== n &&
		e.lp > e.rn &&
		('preserve' === r._attr['xml:space'] && (n = e.readElementContentAsString(!0)),
		(r[i] = n)),
		a(e, r)
	}
}
  
function parseXmlToObject(e) {
var r, n, o
if (e) {
	for (r = new XmlReader(), n = {}, r.reset(), r.setXml(e); r.read(); )
	if (2 !== r.elementType) {
		for (o = {}, o._attr = {}, n[r.name()] = o; r.moveToNextAttribute(); )
		o._attr[r.readAttributeNameAsString()] = r.readContentAsString()
		3 !== r.elementType && a(r, o)
	}
	return n
}
}
