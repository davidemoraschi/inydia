/* =====================================================================================

Canvas-based Rectangle class by Philip Hutchison, March 2009 (http://pipwerks.com).

Copyright (c) 2009 Philip Hutchison
MIT-style license. Full license text can be found at 
http://www.opensource.org/licenses/mit-license.php

Requires MooTools 1.2
Internet Explorer requires excanvas.js since it doesn't natively handle <canvas>

======================================================================================== */


var Rectangle = new Class({
						  
	Implements: [Options],

	//These are defaults that can be overridden by the user.
	options: {
		radius: 0,
		x: 0,
		y: 0,
		width: 0,
		height: 0,
		fill: null,
		easyGradient: null,
		stroke: null,
		canvas: null,
		context: null
	},
	
	initialize: function (options){
		
		if(!options || !options.canvas || !options.width || !options.height || (!options.fill && !options.stroke && !options.easyGradient)){ return false; }
				
		//Pass options to instance
		this.setOptions(options);
		
		//Ensure numbers are rounded
		this.options.width = Math.round(this.options.width);
		this.options.height = Math.round(this.options.height);
		this.options.context = this.options.canvas.getContext('2d') || false;
		
		if(!this.options.context){ return false; }
		
		var opt = this.options; //Shortcut
		
		
		function setGradientFill(obj, height){
			
			var gradientFill = opt.context.createLinearGradient(obj.start_x || 0, obj.start_y || 0, obj.end_x || 0, obj.end_y || height);	
			
			if(obj.stops && $type(obj.stops) === "array"){
				obj.stops.each(function (s){
					if(s.position !== null && s.color){
						gradientFill.addColorStop(s.position, s.color);
					}
				});
			}
			
			return gradientFill;
		
		}


		function configureFill(){
			
			if(opt.easyGradient && opt.easyGradient.length > 0){
				
				opt.fill = { start_x: 0, start_y: 0, end_x: 0, end_y: opt.height, stops: [] };
				
				var stops = opt.easyGradient.length -1;
				
				for(var i=0; i < opt.easyGradient.length; i++){

					var stopPosition = (i === 0) ? 0 : 1 / (stops/i);
					
					opt.fill.stops[i] = {
						color: opt.easyGradient[i],
						position: stopPosition
					};
					
				}

				opt.context.fillStyle = setGradientFill(opt.fill, opt.height);
				
			} else if($type(opt.fill) === "string"){
				
				opt.context.fillStyle = opt.fill;
			
			} else {
				
				opt.context.fillStyle = setGradientFill(opt.fill, opt.height);
			
			}
			
		}
		
		
		function configureStroke(){
			opt.context.lineWidth = Math.round(opt.stroke.width);
			if($type(opt.stroke.color) === "string"){
				opt.context.strokeStyle = opt.stroke.color;
			} else {
				opt.context.strokeStyle = setGradientFill(opt.stroke.color, opt.height);				
			}
		}
		
			
		function draw(){
			
			var edge = { left: opt.x, right: opt.x + opt.width, top: opt.y, bottom: opt.y + opt.height };
			
			opt.context.beginPath();
			
			if(opt.radius){
	
				var isSingleRadius = ($type(opt.radius) === "number");	
				var rad = {
					topLeft: (isSingleRadius) ? opt.radius : (opt.radius.topLeft) ? opt.radius.topLeft : 0,
					topRight: (isSingleRadius) ? opt.radius : (opt.radius.topRight) ? opt.radius.topRight : 0,
					bottomLeft: (isSingleRadius) ? opt.radius : (opt.radius.bottomLeft) ? opt.radius.bottomLeft : 0,
					bottomRight: (isSingleRadius) ? opt.radius : (opt.radius.bottomRight) ? opt.radius.bottomRight : 0
				};
							
				var point = {
					a: opt.y + rad.topLeft,
					b: opt.y + opt.height - rad.bottomLeft,
					c: opt.x + rad.bottomLeft,
					d: opt.x + opt.width - rad.bottomRight,
					e: opt.y + opt.height - rad.bottomRight,
					f: opt.y + rad.topRight,
					g: opt.x + opt.width - rad.topRight,
					h: opt.x + rad.topLeft
				};
		
				opt.context.moveTo(edge.left, point.a);
				opt.context.lineTo(edge.left, point.b);
				opt.context.quadraticCurveTo(edge.left, edge.bottom, point.c, edge.bottom);//Bottom left radius
				opt.context.lineTo(point.d, edge.bottom);
				opt.context.quadraticCurveTo(edge.right, edge.bottom, edge.right, point.e); //Bottom right radius
				opt.context.lineTo(edge.right, point.f);
				opt.context.quadraticCurveTo(edge.right, edge.top, point.g, edge.top); //Top right radius
				opt.context.lineTo(point.h, edge.top);				
				opt.context.quadraticCurveTo(edge.left, edge.top, edge.left, point.a); //Top left radius
		
			} else {
				
				opt.context.moveTo(edge.left, edge.top);
				opt.context.lineTo(edge.left, edge.bottom);
				opt.context.lineTo(edge.right, edge.bottom);
				opt.context.lineTo(edge.right, edge.top);
				opt.context.lineTo(edge.left, edge.top);				
							
			}
			
			if(opt.fill){ opt.context.fill(); }
			if(opt.stroke){ opt.context.stroke(); }
			
		}
				

		if(opt.fill || opt.easyGradient){ 
			configureFill();
		}
		
		if(opt.stroke && opt.stroke.color !== null && opt.stroke.width !== null){
			configureStroke();
		}
		
		draw();
			
	}
		
});