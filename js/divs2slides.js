/**
 * divs2slides.js
 * Ver : 1.3.0
 * update: 10/05/2018
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
                        "style":"width: 90%; padding: 10px; text-align: center;font-size:18px; color: "+data.navTxtColor+";" ////New for Ver: 1.2.1
                    })                
                );
                $("#"+divId+" .slides-toolbar").prepend(
                    $("<img></img>").attr({
                        "id":"slides-next",
                        "class":"slides-nav",
                        "alt":"Next Slide",
                        "src": "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABgAAAAYCAYAAADgdz34AAADZ0lEQVRIiZ2Va2xTdRjGH5xEg0aXeInRqE00pKwfHAkydYMdlXFpiG5LYFFi0iBZFFeLrm50ZrbrYLeyneEFcHMMhxoksGUSTEw01UQz1ukO29quO11vXNZNwBI1gRDN4wfa0q3n1MKT/D697/s8//OenP8BspD2jSGNtmpISEVjcOZmM6sqjcGZm/f2sC2/+tfQpnYvdxwK0NwXpLkvyB2HAixr83C5eUTSGV2GmzZfVjVc+qxlNFb/dYRHh2Psdl7knpNzbDkxy5YTs9xzco7dzov88pdLtHwV5tM1v4W0VUNCVuZ5plO9r3b62O28wF2DURr7zmRk12CUvT9dYLnDw/99Gp3JJW7bL7Pj2zlW9kS4tTucFZU9EX743RwrOibVQ7RVQ0K5w8uG/hm+8knolmjon+Fa+zh12135aQEFtaMh6/HzLN8bSIMkmwajirWFWI6cpc7kcs5fjdFlWNPoYdneAF/qmE6DJP+++i+bBqOK9VQ2fRSkYBvnvJeeZxoeWNcqc2P7tCKp+t79Jys+Dqr2bmyfZkmzjzqTS0wGLH9vNKZ3+KnGQs1evsbaI+dU+/UOP596Z0RKBqy0nOb6VlkVNX3x8yXVmZWW0wQArDCPCIX14yxp8nFt85QimTQ9e5VvHgzP6y+2e24E5O+UNAV1YyyoG2PhBxMsbvDw+UYv1+yeTJJJf135h63fzLC4wcMi6wSfef+6V0HdGJMrKqyfoBKrrG6usrpVzeXoFRr2+RVnn6sfiyUDiqxuabXNTTWU1Pvj76r9q21uFlknBpIBgt0rvtDopRqpisausbIrqNqbQLB7blwZgm1SU9LsoxoJHXf9wTJRVu1L8OJuX0ywSfP/F+taZHFDm0wlSNI+MKNYU2J985Qt7S4SbFKu3uGXMn2h2aB3+KU084T0bXJ+qRiIlXUGeCu8LAaktNUsVKk4qdmyLyS9diDMm2HL/pCzVMxsvijObQByKj+dqjb2RS6bDp9lJoyfhyPbDvheB5ATn034JJUDYDGAOwHcDSAXwIMAHgWwdHPtZ+btnT/01/SMequ7XHJ118hUzUHJ/Van81iFpfddAE8AeATAAwDuBbAEwB0Abo8HXj9xSshd8ZD7ATwM4HEATwJYBkAXRxs3fgzAQwDuA3BPivniuOei/wDo+pj+wU2R5QAAAABJRU5ErkJggg==",
                        "style":"float: right;cursor: pointer;opacity: 0.7;"
                    }).on("click", pptxjslideObj.nextSlide)
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
                        $("<img></img>").attr({
                            "id":"slides-play-pause",
                            "class":"slides-nav-play",
                            "alt":"Play/Pause Slide",
                            "src": "data:image/png;base64,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",
                            "style":"float: left;cursor: pointer;opacity: 0.7;"
                        }).html("<span style='font-size:80%;'>&#x23ef;</span>").bind("click", function(){ //► , ⏯(&#x23ef;)
                            if(data.isSlideMode){
                                pptxjslideObj.startAutoSlide();
                                //TODO : ADD indication that it is in auto slide mode
                            }
                        })
                    );
                }
                $("#"+divId+" .slides-toolbar").prepend(
                    $("<img></img>").attr({
                        "id":"slides-prev",
                        "class":"slides-nav",
                        "alt":"Prev. Slide",
                        "src": "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABgAAAAYCAYAAADgdz34AAADXklEQVRIiZ2Ve2hTZxjGH+tE0bEFdGNM1IBj1PYPKzjr1roeN6tdEW0LWpwMgpMytVncGttF6ZKm2lu0qXdtrdV6wY3ZUsXBYCMbOGrTrce2SZqeNrc5m9bLIjpQxPHsD5PYpMkx+sAPDnzv+zwf3+0ACUipsiiSizuFCL7oVCbSK6tUtVW1SNst5tfZuf2Ui9pWN7Wtbm4/5eK6fQ6mlfzhSfmyy6BUWRQvZJxc3Cm8V/qnR3fey3O/32WT5Q73XhljzeVR1lwe5d4rY2yy3OH3XQGWf+fj+7qewILirryEZ11gsrPlt9vc3eGnuvUvWXZ3+Nlkuc1PG5xM0Vxrea55Yf0AD/w0xqJmHzc1eROiqNnH+h/HuPmoxFSN1RzbfKs1baWxjxVtI9xw2PNSVLSNsMDkYHJxpzAxQGO16C7cYMF+13Op6vCTZMwx/cWbTC/r8UzYVMHQx3UH3VxTPyxLVYef/z76jyRjjufvd3FFpZ2paqtq/OzN2dVOrt43HJfCQ27+bLvP8YpXu6pWYoqmqz0csPCrbjHXNMR4lF34m6P3HjNacj2LdvQEwgFLdNeZUyvF5OzVuxOMQ4rXk1MrcYnuOiMCsox2rqweDLPlpJfDo4/impOMqB9PdpWTGeV9XKztfnqa0nf2Mn1nL5fu6mWmvp9ZFXbWXhrhg4dPZANW7BkIs7zSwawKOzO+7WfIL+0bUQkA+KC8N5BR3s9oVEeGKPkfxg1Yprdxmd42oS9EeIky9f3tHxpsjEfLr7diBsj1ZOptYjhAMNpVH1U6KEdRo5v+QORJkqsXjI5nT4ZgEBUf73EGsqudlCPfLPGi9Z9wgFytYBhQRtzmnOpBwyd1EhPB2D5CknHHV9VIsR+8XNOQKHebEyHXNCQKBjH2D0gwiIq1ZpeY3+Diy5BndgVy66S0mOYh5ZlFxcajHstnx7x8ETYe8Yh55qh1j9KkIEkAJm8+5vxcfdrr05y5QTnUrb57RccHSwBMDvaGfMJKAvAKgKkApgN4HcAbAGYDmF+oa/l6W4Plh9KToq2ksXuwpNEqlTb3OLY2/NK2vuyEFsC7AOYAeBOAAsCrAKYBmBIMxaTgx5RxIa8BmAngLQBzAcwHkAwgNcgCAO8AmAfgbQCzguYzosyT/gcSaJj+/BZ/OQAAAABJRU5ErkJggg==",
                        "style":"float: left;cursor: pointer; opacity: 0.7;",
                    }).bind("click", pptxjslideObj.prevSlide)
                );
                $(".slides-nav, .slides-nav-play").on("mouseover",function(){
                    $(this).css({
                        "opacity":1
                    });
                });
                $(".slides-nav, .slides-nav-play").on("mouseout",function(){
                    $(this).css({
                        "opacity": 0.7
                    });
                });
                if(data.slideCount == 1){
                    $("#"+divId+" #slides-prev").hide();
                }else if(data.slideCount == data.totalSlides){
                    $("#"+divId+" #slides-next").hide();
                }else{
                    $("#"+divId+" #slides-next").show();
                }
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
                    $("#slides-next").show();
            }else{
                if(isLoop){
                    pptxjslideObj.gotoSlide(1);
                }else{
                    $("#slides-next").hide();
                }
            }
            if(data.slideCount > 1){
                $("#slides-prev").show();
            }else{
                $("#slides-prev").hide();
            }
            if(data.slideCount == data.totalSlides && !isLoop){
                $("#slides-next").hide();
            }
            //return this;
        },
        prevSlide: function(){
            var data = pptxjslideObj.data;
            if (data.slideCount > 1){
                pptxjslideObj.gotoSlide(data.slideCount-1);
            }
            if(data.slideCount == 1){
                $("#slides-prev").hide();
            }else{
                $("#slides-prev").show();
            }
            $("#slides-next").show();
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
        var settings = $.extend(true, {
            // These are the defaults.
            first: 1,
            nav: true, /** true,false : show or not nav buttons*/
            showPlayPauseBtn: true, /** true,false */
            navTxtColor: "black", /** color */
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
