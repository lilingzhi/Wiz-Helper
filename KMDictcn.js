(function(f, k) {
	var J, K;

	function L() {
		e.attachEvent && "complete" !== e.readyState || (d.unbindEvent(e, e.addEventListener ? "DOMContentLoaded" : "readystatechange", L), M())
	}
	function M() {
		if (!N) {
			for (var a in v) v[a]();
			N = !0
		}
	}
	function ba(a) {
		return a.replace(/-([a-z])/g, function(a, c) {
			return c.toUpperCase()
		})
	}
	function ca(a) {
		if (!(1 != a.which || t)) {
			var b = a.target;
			if (!h.selfContains(b) && !l.selfContains(b) && (w(), O(), x.selfContains(b))) {
				(b = e.selection) ? (b = window.event, i = {
					pageX: e.body.scrollLeft + j.scrollLeft + b.clientX,
					pageY: e.body.scrollTop + j.scrollTop + b.clientY+8
				}) : i = a;
				b = "";
				a = a.target;
				if ("TEXTAREA" == a.tagName || "INPUT" == a.tagName && "text" == d(a).attr("type")) e.selection ? b = e.selection.createRange().text : f.getSelection && a.selectionStart != k && (b = a.value.substring(a.selectionStart, a.selectionEnd));
				else if ("INPUT" != a.tagName) if (f.getSelection) b = f.getSelection().toString();
				else if (e.getSelection) b = e.getSelection();
				else if (e.selection) b = e.selection.createRange().text;
				20 < b.replace(/[^\x00-\xff]/g, "**").length && (b = "");
				if ((n = P(b)) && g) u || m ? (C(), y()) : (z.css({
					left: i.pageX + "px",
					top: i.pageY + 17 + "px"
				}).show(), D.css({
					left: i.pageX + 8 + "px",
					top: i.pageY + 8 + "px"
				}).show())
			}
		}
	}
	function da(a) {
		27 == a.which && Q()
	}
	function w() {
		z.hide();
		D.hide()
	}
	function ea() {
		w();
		C();
		y(n);
		return !1
	}
	function C() {
		if (!m) {
			var a = i.pageX,
				b = i.pageY + 10,
				c = e.body;
			if (a + 332 > (f.pageXOffset || j.scrollLeft || c.scrollLeft) + (j.clientWidth || c.clientWidth)) a -= 332, b += 8;
			h.css({
				left: a + "px",
				top: b + "px"
			}).show()
		}
		E = !0;
		g && "block" == o.css("display") && (p.appendTo(d("dictHcBottom").dom.firstChild), o.hide())
	}
	function O() {
		!m && E && (h.hide(), E = !1, g || o.append(p).show())
	}
	function Q(a) {
		if (!(a && 1 != a.which)) return m = !1, A.css("background-position", "0 -44px"), w(), O(), F(), !1
	}
	function fa(a) {
		if (1 == a.which) return m = !m, A.css("background-position", "0 -" + (m ? 66 : 44) + "px"), !1
	}
	function ga() {
		q.attr("style", "border:1px solid #A1D9ED;");
		G.attr("checked", r);
		B.attr("checked", u);
		var a = q.offset();
		l.css({
			left: a.left + "px",
			top: a.top - 49 + "px"
		}).show();
		d("dictHcCntLine").show()
	}
	function F(a) {
		clearTimeout(R);
		var b = d("dictHcCntLine");
		if (!a || !(a.relatedTarget == b.dom || a.relatedTarget == l.dom || a.relatedTarget == q.dom)) l.hide(), q.attr("style", null)
	}
	function S() {
		g = !g;
		H("dictHcEnable", g);
		p.css("background-position", "3px -" + (g ? 105 : 127) + "px").html(g ? unescape("%u5DF2%u5F00%u542F%u5212%u8BCD") : unescape("%u5DF2%u5173%u95ED%u5212%u8BCD"));
		return !1
	}
	function ha(a) {
		if ("dictHcBtn1" !== a.target.className) {
			t = !0;
			var b = h.offset();
			J = a.pageX - b.left;
			K = a.pageY - b.top;
			I.show();
			h.attr("class", "dictHcDrag");
			return !1
		}
	}
	function ia(a) {
		if (t) {
			var b = a.pageX - J,
				a = a.pageY - K;
			0 > b && (b = 0);
			0 > a && (a = 0);
			h.css({
				left: b + "px",
				top: a + "px"
			})
		}
	}
	function ja() {
		t && (t = !1, h.attr("class", null), I.hide())
	}
	function P(a) {
		return a.replace(/^\s+|\s+$/g, "").replace(/\s*\r?\n\s*/g, " ")
	}
	function y() {
		T.attr("src", U + "dict.php?skin=default&q=" + encodeURIComponent(n) + "#" + (r ? 1 : 0))
	}
	function H(a, b) {
		d.cookie(a, b, {
			expires: 3650,
			domain: e.domain
		})
	}
	if (!f.dictHc) {
		var V = f.dictHc = {},
			e = f.document,
			j = e.documentElement,
			v = [],
			N = !1,
			W, X, s = document.createElement("div");
		s.style.display = "none";
		s.innerHTML = "<a href='/a' style='color:red;float:left;opacity:.55;'>a</a>";
		s = s.getElementsByTagName("a")[0];
		W = /red/.test(s.getAttribute("style"));
		X = "/a" === s.getAttribute("href");
		var ka = {
			"for": "htmlFor",
			"class": "className",
			readonly: "readOnly",
			maxlength: "maxLength",
			cellspacing: "cellSpacing",
			rowspan: "rowSpan",
			colspan: "colSpan",
			tabindex: "tabIndex",
			usemap: "useMap",
			frameborder: "frameBorder"
		},
			la = /^(?:href|src|style)$/,
			d = function(a) {
				return new d.fn.init(a)
			};
		d.fn = d.prototype = {
			init: function(a) {
				this.dom = "string" === typeof a ? e.getElementById(a) : a.dom ? a.dom : a;
				return this
			},
			attr: function(a, b) {
				if ("object" === typeof a) {
					for (var c in a) this.attr(c, a[c]);
					return this
				}
				a = ka[a] || a;
				c = this.dom;
				var d = b !== k,
					Y = la.test(a);
				if ((a in c || c[a] !== k) && !Y) if (d) c[a] = b;
				else return c[a];
				else if (!W && "style" === a) if (d) c.style.cssText = "" + b;
				else return c.style.cssText;
				else if (d) c.setAttribute(a, "" + b);
				else return !X && Y ? c.getAttribute(a, 2) : c.getAttribute(a);
				return this
			},
			css: function(a, b) {
				if ("object" === typeof a) {
					for (var c in a) this.css(c, a[c]);
					return this
				}
				c = this.dom.style;
				a = ba(a);
				if (b !== k) c[a] = b;
				else return c[a];
				return this
			},
			show: function() {
				return this.css({
					display: "block"
				})
			},
			hide: function() {
				return this.css({
					display: "none"
				})
			},
			html: function(a) {
				var b = this.dom;
				if (a !== k) b.innerHTML = "", b.innerHTML = a;
				else return b.innerHTML;
				return this
			},
			text: function(a) {
				var b = this.dom;
				if (a !== k) this.html("").append(e.createTextNode(a));
				else return d.text([b]);
				return this
			},
			offset: function() {
				var a;
				a = this.dom.getBoundingClientRect();
				var b = e.body;
				return {
					left: a.left + (f.pageXOffset || j.scrollLeft || b.scrollLeft) - (j.clientLeft || b.clientLeft || 0),
					top: a.top + (f.pageYOffset || j.scrollTop || b.scrollTop) - (j.clientTop || b.clientTop || 0)
				}
			},
			bind: function(a, b) {
				var c = this.dom,
					e = {};
				"object" === typeof a ? e = a : e[a] = b;
				for (a in e) d.bindEvent(c, a, e[a]);
				return this
			},
			append: function(a) {
				this.dom.appendChild(a.dom ? a.dom : a);
				return this
			},
			appendTo: function(a) {
				d(a).append(this.dom);
				return this
			},
			contains: function(a) {
				return d.contains(this.dom, a)
			},
			selfContains: function(a) {
				return this.dom === a || d.contains(this.dom, a)
			}
		};
		d.fn.init.prototype = d.fn;
		d.text = function(a) {
			for (var b = "", c, e = 0; a[e]; e++) c = a[e], 3 === c.nodeType || 4 === c.nodeType ? b += c.nodeValue : 8 !== c.nodeType && (b += d.text(c.childNodes));
			return b
		};
		d.contains = function(a, b) {
			return a.contains ? a != b && a.contains(b) : !! (a.compareDocumentPosition(b) & 16)
		};
		d.bindEvent = function(a, b, c) {
			var f = b + c;
			a["e" + f] = c;
			a[f] = function(b) {
				b = b || window.event;
				if (!b.target) b.target = b.srcElement || e;
				if (!b.relatedTarget && b.fromElement) b.relatedTarget = b.fromElement === b.target ? b.toElement : b.fromElement;
				if (null == b.pageX && null != b.clientX) {
					var c = e.documentElement,
						g = e.body;
					b.pageX = b.clientX + (c && c.scrollLeft || g && g.scrollLeft || 0) - (c && c.clientLeft || g && g.clientLeft || 0);
					b.pageY = b.clientY + (c && c.scrollTop || g && g.scrollTop || 0) - (c && c.clientTop || g && g.clientTop || 0)
				}
				if (null == b.which) b.which = null != b.charCode ? b.charCode : null != b.keyCode ? b.keyCode : null;
				if (!b.which && b.button !== k) b.which = b.button & 1 ? 1 : b.button & 2 ? 3 : b.button & 4 ? 2 : 0;
				if ((!/mouse(over|out)/i.test(b.type) || !(b.target !== a && !d.contains(a, b.target) || d.contains(a, b.relatedTarget))) && !1 === a["e" + f](b)) b.preventDefault ? (b.preventDefault(), b.stopPropagation()) : b.returnValue = !1, b.cancelBubble = !0
			};
			a.addEventListener ? a.addEventListener(b, a[f], !1) : a.attachEvent("on" + b, a[f])
		};
		d.unbindEvent = function(a, b, c) {
			c = b + c;
			a.removeEventListener ? a.removeEventListener(b, a[c], !1) : a.detachEvent("on" + b, a[c]);
			a["e" + c] = a[c] = null
		};
		d.bindReady = function(a) {
			if ("complete" === e.readyState) return setTimeout(a, 1);
			v[v.length] = a
		};
		d.bindEvent(e, e.addEventListener ? "DOMContentLoaded" : "readystatechange", L);
		d.bindEvent(f, "load", M);
		d.cookie = function(a, b, c) {
			if ("undefined" != typeof b) {
				var c = c || {},
					d = c.expires;
				null === b && (b = "", d = -1);
				if ("number" === typeof d) {
					var f = d,
						d = new Date;
					d.setDate(d.getDate() + f)
				}
				return e.cookie = [encodeURIComponent(a), "=" + encodeURIComponent(b), d ? "; expires=" + d.toUTCString() : "", c.path ? "; path=" + c.path : "", c.domain ? "; domain=" + c.domain : "", c.secure ? "; secure" : ""].join("")
			}
			return (match = e.cookie.match(RegExp("(^|;| )" + a + "s*=([^;]*)(;|$)", "i"))) ? unescape(match[2]) : null
		};
		var x = e.documentElement,
			U = "http://dict.cn/hc2/",
			Z = "default",
			$ = d(e);
		d(f);
		V.init = function(a) {
			if (a) {
				if (a.area) x = a.area;
				if (a.skin) Z = a.skin
			}
		};
		var ma = function(a, b) {
				if ("string" === typeof a) n = a;
				else {
					var c = d(a);
					n = P(c.text());
					b = c.offset()
				}
				n && (i = {
					pageX: b.left,
					pageY: b.top + 10
				}, w(), C(), y())
			},
			g = !0,
			r = !0,
			u = !1,
			m = !1,
			z, D, h, I, T, o, l, A, aa, p, q, G, B, E = !1,
			t = !1,
			n = "",
			i, R;
		d.bindReady(function() {
			x = d(x);
			d(e.createElement("link")).attr({
				href: U + "skins/" + Z + "/hc.css?v1",
				rel: "stylesheet",
				type: "text/css"
			}).appendTo(e.getElementsByTagName("head")[0]);
			z = d(e.createElement("div")).attr({
				id: "dictHcBtn",
				title: unescape("%u67E5%u8BE2%u5F53%u524D%u9009%u62E9%u8BCD")
			}).css("display", "none").html(unescape("%u67E5%u8BCD%u5178")).appendTo(e.body);
			D = d(e.createElement("div")).attr("id", "dictHcBtnTop").css("display", "none").appendTo(e.body);
			h = d(e.createElement("div"));
			h.attr("id", "dictHc").css("display", "none").html('<div id="dictHcTop" onselectstart="return false"><div><span title="' + unescape("%u9501%u5B9A%u4F4D%u7F6E") + '" id="dictHcLock" class="dictHcBtn1"></span>&nbsp;<span title="' + unescape("%u5173%u95ED") + ' (Esc)" id="dictHcClose" class="dictHcBtn1"></span></div><span id="dictHcTitle">Dict.cn&nbsp;' + unescape("%u6D77%u8BCD") + "&nbsp;-&nbsp;" + unescape("%u5212%u8BCD%u91CA%u4E49") + '</span></div><div id="dictHcDragMask"></div><iframe id="dictHcContent" width="100%" height="100%" frameborder="0"></iframe><div id="dictHcBottom"><div><span id="dictHcEnable" class="dictHcBtn2" title="' + unescape("%u5F00%u542F%u6216%u5173%u95ED%u5212%u8BCD") + '">' + unescape("%u5DF2%u5F00%u542F%u5212%u8BCD") + '</span></div><span id="dictHcSetting" class="dictHcBtn2">' + unescape("%u8BBE%u7F6E") + "</span></div>").appendTo(e.body);
			I = d("dictHcDragMask");
			T = d("dictHcContent");
			A = d("dictHcLock");
			aa = d("dictHcClose");
			p = d("dictHcEnable");
			q = d("dictHcSetting");
			l = d(e.createElement("div"));
			l.attr("id", "dictHcSettingArea").css("display", "none").html('<label for="dictHcSound"><input id="dictHcSound" type="checkbox" value="1" />' + unescape("%u60AC%u505C%u53D1%u97F3") + '</label><label for="dictHcDirect"><input id="dictHcDirect" type="checkbox" value="1" />' + unescape("%u5373%u5212%u5373%u67E5") + '</label><div id="dictHcCntLine" />').appendTo(e.body);
			G = d("dictHcSound");
			B = d("dictHcDirect");
			o = d(e.createElement("div"));
			o.attr({
				id: "dictHcClosetip",
				href: ""
			}).appendTo(e.body);
			$.bind({
				mouseup: ca,
				keydown: da
			});
			z.bind("mouseup", ea);
			A.bind("mouseup", fa);
			aa.bind("mouseup", Q);
			q.bind({
				mouseover: function() {
					R = setTimeout(ga, 200)
				},
				mouseout: F
			});
			l.bind("mouseout", F);
			p.bind("click", S);
			G.bind("change", function() {
				r = d(this).attr("checked") ? !0 : !1;
				if (f._dict_config && f.saveConfig) f._dict_config.ss = r ? _dict_config.ss >> 1 << 1 : f._dict_config.ss | 1, f.saveConfig();
				H("dictHcSound", r);
				y()
			});
			B.bind("change", function() {
				u = B.attr("checked") ? !0 : !1;
				H("dictHcDirect", u)
			});
			d("dictHcTop").bind("mousedown", ha);
			$.bind({
				mousemove: ia,
				mouseup: ja
			});
			var a = d.cookie("dictHcEnable"),
				b = d.cookie("dictHcSound"),
				c = d.cookie("dictHcDirect");
			a && g != ("true" === a) && S();
			b && (r = "true" === b ? !0 : !1);
			c && (u = "true" === c ? !0 : !1);
			g || o.append(p).show();
			V.query = ma
		})
	}
})(window);