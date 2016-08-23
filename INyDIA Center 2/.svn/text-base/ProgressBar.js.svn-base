/* =====================================================================================

Canvas-based ProgressBar class by Philip Hutchison, March 2009 (http://pipwerks.com).

Copyright (c) 2009 Philip Hutchison
MIT-style license. Full license text can be found at 
http://www.opensource.org/licenses/mit-license.php

Requires MooTools 1.2
Internet Explorer requires excanvas.js since it doesn't natively handle <canvas>

======================================================================================== */


var ProgressBar = new Class ({
	
	Implements: [Options],
	
	options: {
		injectInto: null,
		width: 200,
		height: 12,
		border: "1px solid #CCC",
		borderWidth: 1,
		borderColor: "#CCC",
		progressbarGradientColors: ["#C1FC1F", "#393"],
		backgroundGradientColors: ["#DDD", "#FFF"],
		speed: 800
	},
	
	elements: {
		container: null,
		background: null,
		progressbar: null
	},
	
	morph: null,

	initialize: function (options){
		
		if(!options.id || !options.injectInto || !$(options.injectInto) || !options.width || !options.height){ return false; }
		
		if($(options.id)){
			throw new Error("An element with the ID '" +options.id +"' already exists.");   
		}
		
		this.setOptions(options);
		
		var that = this;


		// -- Internal helper functions ----- //

		var makeContainer = function (params){
			
			return new Element("div", {
				id: params.id,
				styles: {
					width: params.width +"px",
					height: params.height +"px",
					border: params.borderWidth +"px solid " +params.borderColor,
					position: "relative"
				}
			}).inject(params.injectInto);
			
		};


		var makeCanvas = function (obj){	
			
			var canvas = $(obj.name);
			
			if(!canvas){
				
				canvas = new Element("canvas", {
					id: obj.id,
					width: obj.width,
					height: obj.height,
					styles: obj.styles || {}
				}).inject(obj.insertInto);
				
				//IE uses VML not Canvas
				var usesVML = (typeof window.G_vmlCanvasManager !== "undefined");
				if(usesVML) { canvas = window.G_vmlCanvasManager.initElement(canvas); }
		
				// Make sure we don't execute when canvas isn't supported
				if(!canvas || !canvas.getContext){ 
					alert("Canvas not supported.\ncanvas.getContext: " +canvas.getContext);
					return false;
				}
				
			}
			
			return canvas;
		
		};
		

		var makeLayer = function (params){
			
			if(!params || !params.width || !params.height || !params.id || !params.insertInto ||!params.colors){ return false; }
			
			var canvas = makeCanvas({
				id: params.id,
				insertInto: params.insertInto,
				width: params.width,
				height: params.height,
				styles: {
					height: params.height,
					display: "block",
					position: "absolute",
					top: 0,
					left: 0,
					"z-index": params.z || "auto"
				}
			});
			
			if(!canvas){ return false; }
			
			new Rectangle({
				canvas: canvas,
				width: params.width,
				height: params.height,
				easyGradient: params.colors
			});
		
			return canvas;
					
		};


		var setMorph = function (el){
			return new Fx.Morph(el, {duration: 1000, transition: Fx.Transitions.Sine.easeOut, link: "chain"})
		};

	
		// -- Perform actions ----- //

		
		that.elements.container = makeContainer({
									id: that.options.id,
									injectInto: that.options.injectInto,
									height: that.options.height, 
									width: that.options.width, 
									borderWidth: that.options.borderWidth,
									borderColor: that.options.borderColor
								  });
		
		
		if(!that.elements.container){
			throw new Error("container element couldn't be created.");
		}

		var adjustedWidth = that.options.width - (that.options.borderWidth *2);
		
		that.elements.background = makeLayer({
			width: that.options.width,
			height: that.options.height,
			id: "progressbarBgnd_canvas",
			insertInto: that.elements.container,
			z: 1,
			colors: that.options.backgroundGradientColors
		});
		
		that.elements.progressbar = makeLayer({
			width: that.options.width,
			height: that.options.height,
			id: "progressbar",
			insertInto: that.elements.container,
			z: 2,
			colors: that.options.progressbarGradientColors
		});

		this.morph = setMorph(that.elements.progressbar).set({width: 0}); //set to zero first		
		return this;
		
	},
	
	set: function (toPercent){
		var px = (this.options.width * toPercent) / 100;		
		this.morph.start({ width: px + "%" });
		return this; //allows chaining
	}
	
});