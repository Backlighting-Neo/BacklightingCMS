//ͼƬ�и�
var ImgCropper = Class.create();
ImgCropper.prototype = {
  //��������,���Ʋ�,ͼƬ��ַ
  initialize: function(container, handle, url, options) {
	this._Container = $(container);//��������
	this._layHandle = $(handle);//���Ʋ�
	this.Url = url;//ͼƬ��ַ
	
	this._layBase = this._Container.appendChild(document.createElement("img"));//�ײ�
	this._layCropper = this._Container.appendChild(document.createElement("img"));//�и��
	this._layCropper.onload = Bind(this, this.SetPos);
	//�������ô�С
	this._tempImg = document.createElement("img");
	this._tempImg.onload = Bind(this, this.SetSize);
	
	this.SetOptions(options);
	
	this.Opacity = Math.round(this.options.Opacity);
	this.Color = this.options.Color;
	this.Scale = !!this.options.Scale;
	this.Ratio = Math.max(this.options.Ratio, 0);
	this.Width = Math.round(this.options.Width);
	this.Height = Math.round(this.options.Height);
	
	//����Ԥ������
	var oPreview = $(this.options.Preview);//Ԥ������
	if(oPreview){
		oPreview.style.position = "relative";
		oPreview.style.overflow = "hidden";
		this.viewWidth = Math.round(this.options.viewWidth);
		this.viewHeight = Math.round(this.options.viewHeight);
		//Ԥ��ͼƬ����
		this._view = oPreview.appendChild(document.createElement("img"));
		this._view.style.position = "absolute";
		this._view.onload = Bind(this, this.SetPreview);
	}
	//�����Ϸ�
	this._drag = new Drag(this._layHandle, { Limit: true, onMove: Bind(this, this.SetPos), Transparent: true });
	//��������
	this.Resize = !!this.options.Resize;
	if(this.Resize){
		var op = this.options, _resize = new Resize(this._layHandle, { Max: true, onResize: Bind(this, this.SetPos) });
		//�������Ŵ�������
		op.RightDown && (_resize.Set(op.RightDown, "right-down"));
		op.LeftDown && (_resize.Set(op.LeftDown, "left-down"));
		op.RightUp && (_resize.Set(op.RightUp, "right-up"));
		op.LeftUp && (_resize.Set(op.LeftUp, "left-up"));
		op.Right && (_resize.Set(op.Right, "right"));
		op.Left && (_resize.Set(op.Left, "left"));
		op.Down && (_resize.Set(op.Down, "down"));
		op.Up && (_resize.Set(op.Up, "up"));
		//��С��Χ����
		this.Min = !!this.options.Min;
		this.minWidth = Math.round(this.options.minWidth);
		this.minHeight = Math.round(this.options.minHeight);
		//�������Ŷ���
		this._resize = _resize;
	}
	//������ʽ
	this._Container.style.position = "relative";
	this._Container.style.overflow = "hidden";
	this._layHandle.style.zIndex = 200;
	this._layCropper.style.zIndex = 100;
	this._layBase.style.position = this._layCropper.style.position = "absolute";
	this._layBase.style.top = this._layBase.style.left = this._layCropper.style.top = this._layCropper.style.left = 0;//����
	//��ʼ������
	this.Init();
  },
  //����Ĭ������
  SetOptions: function(options) {
    this.options = {//Ĭ��ֵ
		Opacity:	50,//͸����(0��100)
		Color:		"",//����ɫ
		Width:		0,//ͼƬ�߶�
		Height:		0,//ͼƬ�߶�
		//���Ŵ�������
		Resize:		false,//�Ƿ���������
		Right:		"",//�ұ����Ŷ���
		Left:		"",//������Ŷ���
		Up:			"",//�ϱ����Ŷ���
		Down:		"",//�±����Ŷ���
		RightDown:	"",//�������Ŷ���
		LeftDown:	"",//�������Ŷ���
		RightUp:	"",//�������Ŷ���
		LeftUp:		"",//�������Ŷ���
		Min:		false,//�Ƿ���С�������(Ϊtrueʱ����min��������)
		minWidth:	50,//��С���
		minHeight:	50,//��С�߶�
		Scale:		false,//�Ƿ񰴱�������
		Ratio:		0,//���ű���(��/��)
		//Ԥ����������
		Preview:	"",//Ԥ������
		viewWidth:	0,//Ԥ�����
		viewHeight:	0//Ԥ���߶�
    };
    Extend(this.options, options || {});
  },
  //��ʼ������
  Init: function() {
	//���ñ���ɫ
	this.Color && (this._Container.style.backgroundColor = this.Color);
	//����ͼƬ
	this._tempImg.src = this._layBase.src = this._layCropper.src = this.Url;
	//����͸��
	if(isIE){
		this._layBase.style.filter = "alpha(opacity:" + this.Opacity + ")";
	} else {
		this._layBase.style.opacity = this.Opacity / 100;
	}
	//����Ԥ������
	this._view && (this._view.src = this.Url);
	//��������
	if(this.Resize){
		with(this._resize){
			Scale = this.Scale; Ratio = this.Ratio; Min = this.Min; minWidth = this.minWidth; minHeight = this.minHeight;
		}
	}
  },
  //�����и���ʽ
  SetPos: function() {
	//ie6��Ⱦbug
	if(isIE6){ with(this._layHandle.style){ zoom = .9; zoom = 1; }; };
	//��ȡλ�ò���
	var p = this.GetPos();
	//���ϷŶ���Ĳ��������и�
	this._layCropper.style.clip = "rect(" + p.Top + "px " + (p.Left + p.Width) + "px " + (p.Top + p.Height) + "px " + p.Left + "px)";
	//����Ԥ��
	this.SetPreview();
  },
  //����Ԥ��Ч��
  SetPreview: function() {
	if(this._view){
		//Ԥ����ʾ�Ŀ�͸�
		var p = this.GetPos(), s = this.GetSize(p.Width, p.Height, this.viewWidth, this.viewHeight), scale = s.Height / p.Height;
		//���������ò���
		var pHeight = this._layBase.height * scale, pWidth = this._layBase.width * scale, pTop = p.Top * scale, pLeft = p.Left * scale;
		//����Ԥ������
		with(this._view.style){
			//������ʽ
			width = pWidth + "px"; height = pHeight + "px"; top = - pTop + "px "; left = - pLeft + "px";
			//�и�Ԥ��ͼ
			clip = "rect(" + pTop + "px " + (pLeft + s.Width) + "px " + (pTop + s.Height) + "px " + pLeft + "px)";
		}
	}
  },
  //����ͼƬ��С
  SetSize: function() {
	var s = this.GetSize(this._tempImg.width, this._tempImg.height, this.Width, this.Height);
	//���õ�ͼ���и�ͼ
	if( s.Width > 600 )
	{
	    var h = Math.round( s.Height * ( 600 / s.Width ) );
		this._layBase.style.width = this._layCropper.style.width = 600 + 'px';
		this._layBase.style.height = this._layCropper.style.height = h + 'px';
		this._Container.style.width = 600 + 'px';
		this._Container.style.height = h + 'px';
		//�����Ϸŷ�Χ
		this._drag.mxRight = 600; this._drag.mxBottom = h;
		//�������ŷ�Χ
		if(this.Resize){ this._resize.mxRight = 600; this._resize.mxBottom = h; }
	}
	else if( s.Height > 500 )
	{
	    var w = Math.round( s.Width * ( 500 / s.Height ) );
		this._layBase.style.width = this._layCropper.style.width = w + 'px';
		this._layBase.style.height = this._layCropper.style.height = 500 + 'px';
		this._Container.style.width = w + 'px';
		this._Container.style.height = 500 + 'px';
		//�����Ϸŷ�Χ
		this._drag.mxRight = w; this._drag.mxBottom = 500;
		//�������ŷ�Χ
		if(this.Resize){ this._resize.mxRight = w; this._resize.mxBottom = 500; }
	}
	else
	{
	    this._layBase.style.width = this._layCropper.style.width = s.Width + "px";
	    this._layBase.style.height = this._layCropper.style.height = s.Height + "px";
		this._Container.style.width = s.Width + 'px';
		this._Container.style.height = s.Height + 'px';
	    //�����Ϸŷ�Χ
	    this._drag.mxRight = s.Width; this._drag.mxBottom = s.Height;
	    //�������ŷ�Χ
	    if(this.Resize){ this._resize.mxRight = s.Width; this._resize.mxBottom = s.Height; }
	}
  },
  //��ȡ��ǰ��ʽ
  GetPos: function() {
	with(this._layHandle){
		return { Top: offsetTop, Left: offsetLeft, Width: offsetWidth, Height: offsetHeight }
	}
  },
  //��ȡ�ߴ�
  GetSize: function(nowWidth, nowHeight, fixWidth, fixHeight) {
	var iWidth = nowWidth, iHeight = nowHeight, scale = iWidth / iHeight;
	//����������
	if(fixHeight){ iWidth = (iHeight = fixHeight) * scale; }
	if(fixWidth && (!fixHeight || iWidth > fixWidth)){ iHeight = (iWidth = fixWidth) / scale; }
	//���سߴ����
	return { Width: iWidth, Height: iHeight }
  }
}