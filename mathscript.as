/*calculated func.*/
/*copyright wxy*/
function HL_4(h1, h2, h3, h4) {
	/*
	| .. h1 .. |
	| .. h2 .. |
	| .. h3 .. |
	| .. h4 .. |
	*/
	lin = h1[0]*HL_3([h2[1], h2[2], h2[3]], [h3[1], h3[2], h3[3]], [h4[1], h4[2], h4[3]])-h2[0]*HL_3([h1[1], h1[2], h1[3]], [h3[1], h3[2], h3[3]], [h4[1], h4[2], h4[3]]);
	lin += h3[0]*HL_3([h1[1], h1[2], h1[3]], [h2[1], h2[2], h2[3]], [h4[1], h4[2], h4[3]])-h4[0]*HL_3([h1[1], h1[2], h1[3]], [h2[1], h2[2], h2[3]], [h3[1], h3[2], h3[3]]);
	return lin;
}
function HL_3(h1, h2, h3) {
	/*
	| .. h1 .. |
	| .. h2 .. |
	| .. h3 .. |
	*/
	return h1[0]*HL_2(h2[1], h2[2], h3[1], h3[2])-h2[0]*HL_2(h1[1], h1[2], h3[1], h3[2])+h3[0]*HL_2(h1[1], h1[2], h2[1], h2[2]);
}
function HL_2(a, b, c, d) {
	/*
	|a  b|
	|c  d|
	*/
	return a*d-b*c;
}
function S(S_a, S_b, S_c) {
	S_s = (S_a+S_b+S_c)/2;
	areas = Math.sqrt(S_s*(S_s-S_a)*(S_s-S_b)*(S_s-S_c));
	return areas;
}
function loadgls(type, widths, heights) {
	_root.scaw = widths;
	_root.scah = heights;
	switch (type) {
	case "gl.3d" :
		gl_x = 0.5;
		gl_y = 0;
		gl_z = 0.5;
		gl_scale = 20;
		_root.createEmptyMovieClip("gl_dragbtn", _root.getNextHighestDepth());
		gl_dragbtn.beginFill(0, 0);
		gl_dragbtn.moveTo(0, 0);
		gl_dragbtn.lineTo(widths, 0);
		gl_dragbtn.lineTo(widths, heights);
		gl_dragbtn.lineTo(0, heights);
		gl_dragbtn.endFill();
		gl_dragbtn.onPress = function() {
			_root._iux = _root._ymouse;
			_root._iuy = _root._xmouse;
			_yuanx = gl_x;
			_yuany = gl_z;
			_root.dong = true;
		};
		gl_dragbtn.onRelease = function() {
			_root.dong = false;
		};
		gl_dragbtn.onEnterFrame = function() {
			if (dong) {
				_dx = -_root._ymouse+_root._iux;
				_dy = -_root._xmouse+_root._iuy;
				_pap = 50;
				gl_x = -_dx/_pap+_yuanx;
				gl_z = _dy/_pap+_yuany;
				gl_Draw();
			}
		};
		break;
		default:
		trace("**错误** mathscript.as: function loadgls(type, widths, heights) \n type='"+type+"'目前不支持。")
	}
}
function gl_L(point1, point2) {
	//length=Sqrt[x^2+y^2+z^2]
	linx = point2[0]-point1[0];
	liny = point2[1]-point1[1];
	linz = point2[2]-point1[2];
	lin = linx*linx+liny*liny+linz*linz;
	return Math.sqrt(lin);
}
function gl_drawl_fx(fx, ts) {
	p1 = fx[0];
	gl_moveTo([p1[0]+fx[1]*ts[0], p1[1]+fx[2]*ts[0], p1[2]+fx[3]*ts[0]]);
	gl_lineTo([p1[0]+fx[1]*ts[1], p1[1]+fx[2]*ts[1], p1[2]+fx[3]*ts[1]]);
}
function gl_Gaverage(p1, p2, p3) {
	return [(p1[0]+p2[0]+p3[0])/3, (p1[1]+p2[1]+p3[1])/3, (p1[2]+p2[2]+p3[2])/3];
}
function gl_divided_point(point1, point2, Lambda) {
	//x=(x1+λx2)/(1+λ);y=(y1+λy2)/(1+λ)
	return [(point1[0]+Lambda*point2[0])/(1+Lambda), (point1[1]+Lambda*point2[1])/(1+Lambda), (point1[2]+Lambda*point2[2])/(1+Lambda)];
}
function gl_solve_line(fx, value, direction) {
	a = fx[0];
	b = fx[1];
	c = fx[2];
	x = value[0];
	y = value[1];
	z = value[2];
	/*return number*/
	if (direction == "∩") {
		//Solve[{x==x0+ t,y==y0+b t,z==z0+c t,a x+b y+c z+1==0},{x,y,z,t}]
		p1 = value[0];
		fgb = value[1];
		fgc = value[2];
		x0 = p1[0];
		y0 = p1[1];
		z0 = p1[2];
		px = (fgb*b*x0-b*y0+fgc*c*x0-c*z0-1)/(a+fgb*b+fgc*c);
		py = (-a*fgb*x0+a*y0-fgb*c*z0-b+fgc*c*y0)/(a+fgb*b+fgc*c);
		pz = (-a*fgc*x0+a*z0+fgb*b*z0-b*fgc*y0-fgc)/(a+fgb*b+fgc*c);
		return [px, py, pz];
		//ABCabcx0y0z0;
	}
	/*if (direction == "h") {
	return (-1-a*value)/b;
	}*/ 
	/*return fx array*/
	if (direction == "⊥") {
		return [value, b/a, c/a];
	}
}
function gl_get_Lfx(p1, p2) {
	gll = gl_L(p1, p2);
	//Solve[{x1==x0+ t,y1==y0+b t,z1==z0+c t}, { b,c,t}]
	fgta = p2[0]-p1[0];
	if (Math.abs(fgta/gll)<0.1) {
		fgtb = p2[1]-p1[1];
		//Solve[{x1==x0+a t,y1==y0+t,z1==z0+c t}, { a,c,t}]
		if (Math.abs(fgtb/gll)<0.1) {
			fgt = p2[2]-p1[2];
			fga = fgta/fgt;
			fgb = fgtb/fgt;
			fgc = 1;
		} else {
			fgb = 1;
			fga = fgta/fgtb;
			fgc = (p2[2]-p1[2])/fgtb;
		}
	} else {
		fga = 1;
		fgb = (p2[1]-p1[1])/fgta;
		fgc = (p2[2]-p1[2])/fgta;
	}
	return [p1, fga, fgb, fgc];
}
function gl_get_Pfx(p1, p2, p3) {
	//Solve[{a*x1 + b*y1 +c z1+ 1 == 0, a*x2 + b*y2 + 1 +c z2== 0, a*x3 + b*y3 + 1 +c z3== 0}, {a, b,c}]
	fga = -(p2[1]*p1[2]-p3[1]*p1[2]-p1[1]*p2[2]+p3[1]*p2[2]+p1[1]*p3[2]-p2[1]*p3[2])/(p3[0]*p2[1]*p1[2]-p2[0]*p3[1]*p1[2]-p3[0]*p1[1]*p2[2]+p1[0]*p3[1]*p2[2]+p2[0]*p1[1]*p3[2]-p1[0]*p2[1]*p3[2]);
	fgb = -(-p2[0]*p1[2]+p3[0]*p1[2]+p1[0]*p2[2]-p3[0]*p2[2]-p1[0]*p3[2]+p2[0]*p3[2])/(p3[0]*p2[1]*p1[2]-p2[0]*p3[1]*p1[2]-p3[0]*p1[1]*p2[2]+p1[0]*p3[1]*p2[2]+p2[0]*p1[1]*p3[2]-p1[0]*p2[1]*p3[2]);
	fgc = -(-p2[0]*p1[1]+p3[0]*p1[1]+p1[0]*p2[1]-p3[0]*p2[1]-p1[0]*p3[1]+p2[0]*p3[1])/(-p3[0]*p2[1]*p1[2]+p2[0]*p3[1]*p1[2]+p3[0]*p1[1]*p2[2]-p1[0]*p3[1]*p2[2]-p2[0]*p1[1]*p3[2]+p1[0]*p2[1]*p3[2]);
	return [fga, fgb, fgc];
}
function gl_DymoveTo(P, l) {
	gl_moveTo([0+P[0], 0+P[1], -l+P[2]]);
	gl_lineTo([0+P[0], 0+P[1], l+P[2]]);
	gl_moveTo([0+P[0], -l+P[1], 0+P[2]]);
	gl_lineTo([0+P[0], l+P[1], 0+P[2]]);
	gl_moveTo([-l+P[0], 0+P[1], 0+P[2]]);
	gl_lineTo([l+P[0], 0+P[1], 0+P[2]]);
}
function solveNerror(num) {
	if (Math.abs(num)<0.00001) {
		num = 0;
	}
	return num;
}
function gl_tetrohedro(P1, P2, P3, P4) {
	gl_moveTo([P1[0], P1[1], P1[2]]);
	gl_lineTo([P2[0], P2[1], P2[2]]);
	gl_lineTo([P3[0], P3[1], P3[2]]);
	gl_lineTo([P1[0], P1[1], P1[2]]);
	gl_lineTo([P4[0], P4[1], P4[2]]);
	gl_lineTo([P2[0], P2[1], P2[2]]);
	gl_moveTo([P4[0], P4[1], P4[2]]);
	gl_lineTo([P3[0], P3[1], P3[2]]);
}
function gl_boxTo(point3d1, point3d2) {
	gl_moveTo([point3d1[0], point3d1[1], point3d1[2]]);
	gl_lineTo([point3d1[0], point3d1[1], point3d2[2]]);
	gl_lineTo([point3d1[0], point3d2[1], point3d2[2]]);
	gl_lineTo([point3d1[0], point3d2[1], point3d1[2]]);
	gl_lineTo([point3d1[0], point3d1[1], point3d1[2]]);
	gl_lineTo([point3d2[0], point3d1[1], point3d1[2]]);
	gl_lineTo([point3d2[0], point3d1[1], point3d2[2]]);
	gl_lineTo([point3d2[0], point3d2[1], point3d2[2]]);
	gl_lineTo([point3d2[0], point3d2[1], point3d1[2]]);
	gl_lineTo([point3d2[0], point3d1[1], point3d1[2]]);
	gl_moveTo([point3d1[0], point3d2[1], point3d2[2]]);
	gl_lineTo([point3d2[0], point3d2[1], point3d2[2]]);
	gl_moveTo([point3d1[0], point3d1[1], point3d2[2]]);
	gl_lineTo([point3d2[0], point3d1[1], point3d2[2]]);
	gl_moveTo([point3d1[0], point3d2[1], point3d1[2]]);
	gl_lineTo([point3d2[0], point3d2[1], point3d1[2]]);
}
function gl_moveTo(point3ds) {
	gl_P = gl_GetPoint3d(point3ds);
	//traces(gl_P)
	//if (obj == undefined) {
	_root.moveTo(gl_P[0]*_root.gl_scale+_root.scaw/2, gl_P[1]*_root.gl_scale+_root.scah/2);
	/*} else {
	obj.moveTo(gl_P[0], gl_P[1]);
	}*/
}
function gl_lineTo(point3ds) {
	gl_P = gl_GetPoint3d(point3ds);
	//traces(gl_P)
	//if (obj == undefined) {
	_root.lineTo(gl_P[0]*_root.gl_scale+_root.scaw/2, gl_P[1]*_root.gl_scale+_root.scah/2);
	/*} else {
	obj.lineTo(gl_P[0], gl_P[1]);
	}*/
}
function gl_GetPoint3d(ttpoint3d) {
	/**Key code start to dell point error**/
	point3d = [ttpoint3d[0], ttpoint3d[1], ttpoint3d[2]];
	/**Key code end to dell point error**/
	_point3d = rotate(_root.gl_z, [point3d[0], point3d[1]], [0, 0]);
	point3d[0] = _point3d[0];
	point3d[1] = _point3d[1];
	//z-xy
	_point3d = rotate(_root.gl_x, [point3d[1], point3d[2]], [0, 0]);
	point3d[1] = _point3d[0];
	point3d[2] = _point3d[1];
	//x-yz*/
	_point3d = rotate(_root.gl_y, [point3d[0], point3d[2]], [0, 0]);
	point3d[0] = _point3d[0];
	point3d[2] = _point3d[1];
	//y-xz
	return [point3d[0], point3d[2]];
}
function rotate(sita, point1, point2) {
	_an_ = atan2(point2, point1)+sita;
	_r_ = L(point1, point2);
	_NX_ = solveNerror(Math.cos(_an_)*_r_+point2[0]);
	_NY_ = solveNerror(Math.sin(_an_)*_r_+point2[1]);
	//trace(_NX_+","+_NY_)
	return [_NX_, _NY_];
}
function degreetorad(sita) {
	return (sita/180)*Math.PI;
}
function radtodegree(sita) {
	return (sita/Math.PI)*180;
}
function Draw_fx(fx, obj) {
	if (obj == undefined) {
		if (Math.abs(tg(fx))<=1) {
			_root.moveTo(-500, solve_point(fx, -500, "y"));
			_root.lineTo(1000, solve_point(fx, 1000, "y"));
		} else {
			_root.moveTo(solve_point(fx, -500, "x"), -500);
			_root.lineTo(solve_point(fx, 1000, "x"), 1000);
		}
	} else {
		if (Math.abs(tg(fx))<=1) {
			obj.moveTo(-500, solve_point(fx, -500, "y"));
			obj.lineTo(1000, solve_point(fx, 1000, "y"));
		} else {
			obj.moveTo(solve_point(fx, -500, "x"), -500);
			obj.lineTo(solve_point(fx, 1000, "x"), 1000);
		}
	}
}
/*mathematical func.*/
function circleTo(O, R, obj) {
	c_step = 140;
	if (obj == undefined) {
		_root.moveTo(O[0]+R, O[1]);
		for (c_angle=0; c_angle<=2*Math.PI; c_angle += 2/c_step*Math.PI) {
			_root.lineTo(R*Math.cos(c_angle)+O[0], R*Math.sin(c_angle)+O[1]);
		}
		_root.lineTo(O[0]+R, O[1]);
	} else {
		obj.moveTo(O[0]+R, O[1]);
		for (c_angle=0; c_angle<=2*Math.PI; c_angle += 2/c_step*Math.PI) {
			obj.lineTo(R*Math.cos(c_angle)+O[0], R*Math.sin(c_angle)+O[1]);
		}
		obj.lineTo(O[0]+R, O[1]);
	}
}
function P(obj) {
	return [obj._x, obj._y];
}
function divided_point(point1, point2, Lambda) {
	//x=(x1+λx2)/(1+λ);y=(y1+λy2)/(1+λ)
	return [(point1[0]+Lambda*point2[0])/(1+Lambda), (point1[1]+Lambda*point2[1])/(1+Lambda)];
}
function atan2(ponit1, ponit2) {
	return Math.atan2(ponit2[1]-ponit1[1], ponit2[0]-ponit1[0]);
}
function tg(fx) {
	//ax==by -> y=a/b*x -> k=a/b
	return (fx[0]/fx[1]);
}
function L(point1, point2) {
	//length=Sqrt[x^2+y^2]
	return Math.sqrt((point2[0]-point1[0])*(point2[0]-point1[0])+(point2[1]-point1[1])*(point2[1]-point1[1]));
}
function get_fx(point1, point2) {
	//Solve[{a*x1 + b*y1 + 1 == 0, a*x2 + b*y2 + 1 == 0}, {a, b}]
	x1 = point1[0];
	x2 = point2[0];
	y1 = point1[1];
	y2 = point2[1];
	a = -((-y1+y2)/(-x2*y1+x1*y2));
	b = -((x1-x2)/(-x2*y1+x1*y2));
	return [a, b];
}
function get_common_point(fx1, fx2) {
	//Solve[{a1*x + b1*y + 1 == 0, a2*x + b2*y + 1 == 0}, {x, y}]
	a1 = fx1[0];
	a2 = fx2[0];
	b1 = fx1[1];
	b2 = fx2[1];
	x = -((-b1+b2)/(-a2*b1+a1*b2));
	y = -((a1-a2)/(-a2*b1+a1*b2));
	return [x, y];
}
function solve_point(fx, value, direction) {
	//Solve[{a*x + b*y + 1 == 0},{x}] or Solve[{a*x + b*y + 1 == 0},{y}]
	a = fx[0];
	b = fx[1];
	x = value[0];
	y = value[1];
	/*return number*/
	if (direction == "x") {
		return (-1-b*value)/a;
	}
	if (direction == "y") {
		return (-1-a*value)/b;
	}
	/*return fx array*/ 
	if (direction == "∥") {
		return [-(a/(a*x+b*y)), -(b/(a*x+b*y))];
	}
	if (direction == "⊥") {
		return [-(b/(b*x-a*y)), -(a/(-b*x+a*y))];
	}
}
