/**
 * divs2slides.js
 * Ver : 1.2.1
 * Author: meshesha , https://github.com/meshesha
 * LICENSE: MIT
 * url:https://github.com/meshesha/divs2slides
 */
(function( $ ){
    var pptxjslideObj = {
        init: function(){
            var data = pptxjslideObj.data;
            var divId = data.divId;
            var isInit = data.isInit;
            $("#"+divId+" .slide").hide();        
            if(data.slctdBgClr != false){
                var preBgClr = $(document.body).css("background-color");
                data.prevBgColor = preBgClr;
                $(document.body).css("background-color",data.slctdBgClr)
            }
            if (data.nav && !isInit){
                data.isInit = true;
                // Create navigators 
                $("#"+divId).prepend(
                    $("<div></div>").attr({
                        "class":"slides-toolbar",
                        "style":"width: 100%; padding: 10px; text-align: center; color: "+data.navTxtColor+";" ////New for Ver: 1.2.1
                    })                
                );
                $("#"+divId+" .slides-toolbar").prepend(
                    $("<button></button>").attr({
                        "id":"slides-next",
                        "class":"slides-nav",
                        "alt":"Next Slide",
                        "style":"float: right; "
                    }).html(data.navNextTxt).on("click", pptxjslideObj.nextSlide)
                );
                if(data.showTotalSlideNum){
                    $("#"+divId+" .slides-toolbar").prepend(
                        $("<span></span>").attr({
                            "id":"slides-total-slides-num"
                        }).html(data.totalSlides)
                    );
                }
                if(data.showSlideNum && data.showTotalSlideNum){
                    $("#"+divId+" .slides-toolbar").prepend(
                        $("<span></span>").attr({
                            "id":"slides-slides-num-separator"
                        }).html(" / ")
                    );
        
                }
                if(data.showSlideNum){
                    $("#"+divId+" .slides-toolbar").prepend(
                        $("<span></span>").attr({
                            "id":"slides-slide-num"
                        }).html(data.slideCount)
                    );
                }
                if(data.showPlayPauseBtn){
                    $("#"+divId+" .slides-toolbar").prepend(
                        $("<button></button>").attr({
                            "id":"slides-prev",
                            "alt":"Play/Pause Slide",
                            "style":"float: left;"
                        }).html("<span style='font-size:80%;'>&#x23ef;</span>").bind("click", function(){ //► , ⏯(&#x23ef;)
                            if(data.isSlideMode){
                                pptxjslideObj.startAutoSlide();
                                //TODO : ADD indication that it is in auto slide mode
                            }
                        })
                    );
                }
                $("#"+divId+" .slides-toolbar").prepend(
                    $("<button></button>").attr({
                        "id":"slides-prev",
                        "class":"slides-nav",
                        "alt":"Prev. Slide",
                        "style":"float: left;"
                    }).html(data.navPrevTxt).bind("click", pptxjslideObj.prevSlide)
                );
            }else{
                $("#"+divId+" .slides-toolbar").show();
                data.isEnbleNextBtn = true;
                data.isEnblePrevBtn = true;
            }
            // Go to first slide
            pptxjslideObj.gotoSlide(1);
        },
        nextSlide: function(){
            var data = pptxjslideObj.data;
            var isLoop = data.isLoop;
            if (data.slideCount < data.totalSlides){
                    pptxjslideObj.gotoSlide(data.slideCount+1);
            }else{
                if(isLoop){
                    pptxjslideObj.gotoSlide(1);
                }
            }
            //return this;
        },
        prevSlide: function(){
            var data = pptxjslideObj.data;
            if (data.slideCount > 1){
                pptxjslideObj.gotoSlide(data.slideCount-1);
                //pptxjslideObj.data.slideCount--;
            }
            return this;
        },
        gotoSlide: function(idx){
            var index = idx - 1;
            var data = pptxjslideObj.data;
            var slides = data.slides;
            var prevSlidNum = data.prevSlide;
            var transType = data.transition; /*"slid","fade","default" */
            if(transType=="random"){
                var tType = ["","default","fade","slid"];
                var randomNum = Math.floor(Math.random() * 3) + 1; //random number between 1 to 3
                transType = tType[randomNum];
            }
            var transTime = 1000*(data.transitionTime);
            if (slides[index]){
                var nextSlide = $(slides[index]);
                if ($(slides[prevSlidNum]).is(":visible")){ //remove "index >= 1 &&" bugFix to ver. 1.2.1
                    if(transType=="default"){
                        $(slides[prevSlidNum]).hide(transTime);
                    }else if(transType=="fade"){
                        $(slides[prevSlidNum]).fadeOut(transTime);
                    }else if(transType=="slid"){
                        $(slides[prevSlidNum]).slideUp(transTime);
                    }
                }
                if(transType=="default"){
                    nextSlide.show(transTime); 
                }else if(transType=="fade"){
                    nextSlide.fadeIn(transTime);
                }else if(transType=="slid"){
                    nextSlide.slideDown(transTime);
                }
                data.prevSlide = index;
                pptxjslideObj.data.slideCount = idx;
                $("#slides-slide-num").html(idx);
            }
            return this;
        },
        keyDown: function(event){
            event.preventDefault();
            var key = event.keyCode;
            //console.log(key);
            var data = pptxjslideObj.data;
            switch(key){
                case(37): // Left arrow
                case(8): // Backspace
                    if(data.isSlideMode && data.isEnblePrevBtn){
                        pptxjslideObj.prevSlide();
                    }
                    break;
                case(39): // Right arrow
                case(32): // Space 
                case(13): // Enter 
                    if(data.isSlideMode  && data.isEnbleNextBtn){
                        pptxjslideObj.nextSlide();
                    }
                    break; 
                case(46): //Delete
                    //if in auto mode , stop auto mode TODO
                    if(data.isSlideMode){
                        var div_id = data.divId;
                        $("#"+div_id+" .slide").hide();
                        pptxjslideObj.gotoSlide(1);               //bugFix to ver. 1.2.1
                    }
                    break;
                case(27): //Esc
                    if(data.isSlideMode){
                        pptxjslideObj.closeSileMode();
                        data.isSlideMode = false;
                    }
                    break;
                case(116): //F5
                    if(!data.isSlideMode){
                        pptxjslideObj.startSlideMode();
                        data.isSlideMode = true;
                        if(data.isAutoSlideMode || data.isLoopMode){
                            clearInterval(data.loopIntrval);
                            data.isAutoSlideMode = false;
                            data.isLoopMode = false;
                        }
                        
                    }
                    break;
                case(113): // F2
                    if(data.isSlideMode){
                        pptxjslideObj.fullscreen();
                    }
                    break;
                case(119): // F8
                    if(data.isSlideMode){
                        pptxjslideObj.startAutoSlide();
                        //TODO : ADD indication that it is in auto slide mode
                    }
                break;
            }
            return true;
        },
        startSlideMode: function(){
            pptxjslideObj.init();
        },
        closeSileMode: function(){
            var data = pptxjslideObj.data;
            data.isSlideMode = false;
            var div_id= data.divId;
            $("#"+div_id+" .slides-toolbar").hide();
            $("#"+div_id+" .slide").show();
            $(document.body).css("background-color",pptxjslideObj.data.prevBgColor);
            if(data.isLoopMode){
                clearInterval(data.loopIntrval);
                data.isLoopMode = false;
            }
            
        },
        startAutoSlide: function(){
            var data = pptxjslideObj.data;
            var isAutoSlideOption = data.timeBetweenSlides
            var isAutoSlideMode = data.isAutoSlideMode;
            if(!isAutoSlideMode && isAutoSlideOption !== false){
                data.isAutoSlideMode = true;
                //var isLoopOption = data.isLoop;
                var isStrtLoop =  data.isLoopMode;
                //hide and disable next and prev btn
                if(data.nav){
                    var div_Id = data.divId;
                    $("#"+div_Id+" .slides-toolbar .slides-nav").hide();
                }
                data.isEnbleNextBtn = false;
                data.isEnblePrevBtn = false;
                ///////////////////////////////
                
                var t = isAutoSlideOption + data.transitionTime;
                
                var slideNums = data.totalSlides;
                var isRandomSlide = data.randomAutoSlide;
                
                if(!isStrtLoop){
                    var timeBtweenSlides = t*1000; //milisecons
                    data.isLoopMode = true;
                    data.loopIntrval = setInterval(function(){
                        if(isRandomSlide){
                            var randomSlideNum = Math.floor(Math.random() * slideNums) + 1;
                            pptxjslideObj.gotoSlide(randomSlideNum);
                        }else{
                            pptxjslideObj.nextSlide();
                        }
                    }, timeBtweenSlides);
                }else{
                    clearInterval(data.loopIntrval);
                    data.isLoopMode = false;                
                }
            }else{
                clearInterval(data.loopIntrval);
                data.isAutoSlideMode = false;
                data.isLoopMode = false;
                //show and enable next and prev btn
                if(data.nav){
                    var div_Id = data.divId;
                    $("#"+div_Id+" .slides-toolbar .slides-nav").show();
                }
                data.isEnbleNextBtn = true;
                data.isEnblePrevBtn = true;    
            }
        },
        fullscreen: function(){

            if (!document.fullscreenElement &&    
                !document.mozFullScreenElement && !document.webkitFullscreenElement && !document.msFullscreenElement ) {  // current working methods
              if (document.documentElement.requestFullscreen) {
                document.documentElement.requestFullscreen();
              } else if (document.documentElement.msRequestFullscreen) {
                document.documentElement.msRequestFullscreen();
              } else if (document.documentElement.mozRequestFullScreen) {
                document.documentElement.mozRequestFullScreen();
              } else if (document.documentElement.webkitRequestFullscreen) {
                document.documentElement.webkitRequestFullscreen(Element.ALLOW_KEYBOARD_INPUT);
              }
            } else {
              if (document.exitFullscreen) {
                document.exitFullscreen();
              } else if (document.msExitFullscreen) {
                document.msExitFullscreen();
              } else if (document.mozCancelFullScreen) {
                document.mozCancelFullScreen();
              } else if (document.webkitExitFullscreen) {
                document.webkitExitFullscreen();
              }
            }
            
        }

    };
    $.fn.divs2slides = function( options ) {
        var target = $(this);
        var divId = target.attr("id");
        var slides = target.children();
        var totalSlides = slides.length-1;
        var prevBgColor;
        var settings = $.extend({
            // These are the defaults.
            first: 1,
            nav: true, /** true,false : show or not nav buttons*/
            showPlayPauseBtn: true, /** true,false */
            navTxtColor: "black", /** color */
            navNextTxt:"&#8250;",
            navPrevTxt: "&#8249;",
            keyBoardShortCut: true, /** true,false */
            showSlideNum: true, /** true,false */
            showTotalSlideNum: true, /** true,false */
            autoSlide:false, /** false or seconds (the pause time between slides) , F8 to active(condition: keyBoardShortCut: true) */
            randomAutoSlide: false, /** true,false ,(condition: autoSlide:true */ 
            loop: false,  /** true,false */
            background: false, /** false or color*/
            transition: "default", /** transition type: "slid","fade","default","random" , to show transition efects :transitionTime > 0.5 */
            transitionTime: 1 /** transition time in seconds */
        }, options );
        var slideCount = settings.first
        pptxjslideObj.data = {
            nav: settings.nav,
            navTxtColor: settings.navTxtColor,
            navNextTxt: settings.navNextTxt,
            navPrevTxt: settings.navPrevTxt,
            showPlayPauseBtn: settings.showPlayPauseBtn,
            showSlideNum: settings.showSlideNum,
            showTotalSlideNum: settings.showTotalSlideNum,
            target: target,
            divId: divId,
            slides:slides,
            isSlideMode: true,
            totalSlides:totalSlides,
            slideCount: slideCount,
            prevSlide: 0,
            transition: settings.transition,
            transitionTime: settings.transitionTime,
            slctdBgClr: settings.background,
            prevBgColor: prevBgColor,
            timeBetweenSlides: settings.autoSlide,
            isLoop: settings.loop,
            isLoopMode: false,
            isAutoSlideMode: false,
            randomAutoSlide: settings.randomAutoSlide,
            isEnbleNextBtn: true,
            isEnblePrevBtn: true,
            isInit: false
        }

        // Keyboard shortcuts
        if (settings.keyBoardShortCut){
            $(document).bind("keydown",pptxjslideObj.keyDown);
        }
        pptxjslideObj.init();
    }
})(jQuery);
