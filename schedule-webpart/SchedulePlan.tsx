import * as ICS from 'ics-js';
import * as download from 'downloadjs';

import * as strings from 'ScheduleWebPartStrings';

export default class SchedulePlan {
    private transitionEnd:any;
    private transitionsSupported:any;

    private element:any;
    private timeline:any;
    private timelineItems:any;
    private timelineItemsNumber:any;
    private timelineStart:any;
    private timelineUnitDuration:any;
    private eventsWrapper:any;
    private eventsGroup:any;
    private singleEvents:any;
    private eventSlotHeight:any;
    private modal:any;
    private modalHeader:any;
    private modalHeaderBg:any;
    private modalBody:any;
    private modalBodyBg:any;
    private modalMaxWidth:any;
    private modalMaxHeight:any;
    private animating:any;

    constructor(element) {
        this.transitionEnd = 'webkitTransitionEnd otransitionend oTransitionEnd msTransitionEnd transitionend';
        this.transitionsSupported = ( $('.csstransitions').length > 0 );

        if(!this.transitionsSupported) {
            this.transitionEnd = 'noTransition';
        }

        this.element = element;
        this.timeline = this.element.find('.timeline');
        this.timelineItems = this.timeline.find('li');
        this.timelineItemsNumber = this.timelineItems.length;
        this.timelineStart = this.getScheduleTimestamp(this.timelineItems.eq(0).text());

        this.timelineUnitDuration = this.getScheduleTimestamp(this.timelineItems.eq(1).text()) - this.getScheduleTimestamp(this.timelineItems.eq(0).text());

        this.eventsWrapper = this.element.find('.events');
        this.eventsGroup = this.eventsWrapper.find('.events-group');
        this.singleEvents = this.eventsGroup.find('.single-event');
        this.eventSlotHeight = this.eventsGroup.eq(0).children('.top-info').outerHeight();

        this.modal = this.element.find('.event-modal');
        this.modalHeader = this.modal.find('.header');
        this.modalHeaderBg = this.modal.find('.header-bg');
        this.modalBody = this.modal.find('.body');
        this.modalBodyBg = this.modal.find('.body-bg');
        this.modalMaxWidth = 800;
        this.modalMaxHeight = 480;

        this.animating = false;

        this.initSchedule();
    }

    private initSchedule() {
        this.scheduleReset();
        this.initEvents();
      };

    private scheduleReset() {
        var mq = this.mq();
        if( mq == 'desktop' && !this.element.hasClass('js-full') ) {
          //in this case you are on a desktop version (first load or resize from mobile)
          this.eventSlotHeight = this.eventsGroup.eq(0).children('.top-info').outerHeight();
          this.element.addClass('js-full');
          this.placeEvents();
        } else if(  mq == 'mobile' && this.element.hasClass('js-full') ) {
          //in this case you are on a mobile version (first load or resize from desktop)
          this.element.removeClass('js-full loading');
          this.eventsGroup.children('ul').add(this.singleEvents).removeAttr('style');
          this.eventsWrapper.children('.grid-line').remove();
        } else if( mq == 'desktop' && this.element.hasClass('modal-is-open')){
          //on a mobile version with modal open - need to resize/move modal window
          this.checkEventModal();
          this.element.removeClass('loading');
        } else {
          this.element.removeClass('loading');
        }
    }

    private initEvents() {
        var self = this;

        this.singleEvents.each(function(){
          //create the .event-date element for each event
          var durationLabel = '<span class="event-date">'+$(this).data('start')+' - '+$(this).data('end')+'</span>';
          $(this).children('a').prepend($(durationLabel));

          //detect click on the event and open the modal
          $(this).on('click', 'a', function(event){
            event.preventDefault();
            if( !self.animating ) self.openModal($(this));
          });
        });

        //close modal window
        this.element.on('click', '.close', (event) => {
          event.preventDefault();
          if( !self.animating ) self.closeModal(self.eventsGroup.find('.selected-event'));
        });
        this.element.on('click', '.cover-layer', (event) => {
          if( !self.animating && self.element.hasClass('modal-is-open') ) self.closeModal(self.eventsGroup.find('.selected-event'));
        });
    }

    private placeEvents() {
        var self = this;
        this.singleEvents.each(function(){
          //place each event in the grid -> need to set top position and height
          var start = self.getScheduleTimestamp($(this).attr('data-start')),
            duration = self.getScheduleTimestamp($(this).attr('data-end')) - start;

          var eventTop = self.eventSlotHeight*(start - self.timelineStart)/self.timelineUnitDuration,
            eventHeight = self.eventSlotHeight*duration/self.timelineUnitDuration;

          $(this).css({
            top: (eventTop -1) +'px',
            height: (eventHeight+1)+'px'
          });
        });

        this.element.removeClass('loading');
    }

    public openSurvey(sID, wID, iID) {
      window.location.href = "./" + strings.SurveyUrl + ".aspx?sID=" + sID + "&wID=" + wID + "&iID=" + iID;
    }

    public async downloadICS(id, title, location, from, to) {

      var cal = new ICS.VCALENDAR();
      cal.addProp('VERSION', 1);
      cal.addProp('CALSCALE', 'GREGORIAN');
      cal.addProp('PRODID', 'Test Company d.o.o.');

      const event = new ICS.VEVENT();
      event.addProp('UID');
      event.addProp('SUMMARY', title.replace(/–/g, '-').replace(/’/g, '\''));
      event.addProp('DTSTAMP', from);
      event.addProp('DTSTART', from);
      event.addProp('DTEND', to);
      event.addProp('LOCATION', location);

      cal.addComponent(event);

      download('data:text/plain;base64,' + cal.toBase64(), title + ".ics", "text/plain");
    }

    private openModal(event) {
        var self = this;
        var mq = self.mq();
        this.animating = true;

        //update event name and time
        this.modalHeader.find('.event-name').text(event.find('.event-name').text());
        this.modalHeader.find('.event-author').text(event.find('.event-author').text());
        this.modalHeader.find('.event-date').text(event.find('.event-date').text());

        this.modalHeader.find('.event-ics').css("cursor", "pointer");
        this.modalHeader.find('.event-ics').on('click', (e) => {
          e.preventDefault();

          var title = event.find('.event-name').text();
          var ID = event.parent().attr('data-id');
          var location = event.parent().attr('data-location');
          var dtFrom = event.parent().attr('data-dtFrom');
          var dtTo = event.parent().attr('data-dtTo');

          this.downloadICS(ID, title, location, dtFrom, dtTo);
        });

        this.modalHeader.find('.event-survey').css("cursor", "pointer");
        this.modalHeader.find('.event-survey').on('click', (e) => {
          e.preventDefault();

          var sID = event.parent().attr('data-sid');
          var wID = event.parent().attr('data-wid');
          var iID = event.parent().attr('data-iid');

          this.openSurvey(sID, wID, iID);
        });

        this.modal.attr('data-event', event.parent().attr('data-event'));

        this.modalBody.find('.event-info #ifmReport').load(function() {
          $('.event-info #ifmReport').contents().find('#suiteBarDelta, #s4-ribbonrow, #titlerow, #social-div, .footer-wrapper, .cc-window.cc-floating.cc-type-info.cc-theme-classic').css("display", "none");
          $('.event-info #ifmReport').contents().find('#s4-workspace').css("overflow-x", "hidden");
          $('.event-info #ifmReport').contents().find('#contentBox').css("width", "100%");
          $('.event-info #ifmReport').contents().find('#ms-belltown-table, .contentwrapper > .ms-table.ms-fullWidth').css("table-layout", "fixed");


          $('.event-info #ifmReport').contents().find('#s4-bodyContainer > div.ms-table').css("background", "transparent");
          $('.event-info #ifmReport').contents().find('.contentwrapper').css("margin", "20px 0 60px 0");
          $('.event-info #ifmReport').contents().find('#contentBox').css("padding", "0 10px");
          $('.event-info #ifmReport').contents().find('#SessionOne .headerclass1, #SessionOne .headerclass2').css({
            'font-size' : '32px',
            'line-height' : '40px'
          });
          $('.event-info #ifmReport').contents().find('#SessionTwo').css("padding", "0px");
          $('.event-info #ifmReport').contents().find('#SessionTwo .headerclass1').css({
            'font-size' : '24px',
            'line-height' : '30px'
          });
          $('.event-info #ifmReport').contents().find('#SessionTwo .predavanje').css("min-width", "200px");
          $('.event-info #ifmReport').contents().find('#SessionTwo .predavatelji, #SessionTwo .predavanje').css("display", "inline-block");
          $('.event-info #ifmReport').contents().find('#SessionTwo .predavatelji').css("width", "170px");
          $('.event-info #ifmReport').contents().find('#SessionTwo .predavanje').css("width", "calc(100% - 180px)");
          $('.event-info #ifmReport').contents().find('#SessionTwo .predavatelji .predavatelj').css({
            'display' : 'block',
            'width' : '150px',
            'height' : '150px'
          });
          $('.event-info #ifmReport').contents().find('#SessionTwo .predavatelji .predavatelj .slika').css("height", "100px");
          $('.event-info #ifmReport').contents().find('#SessionTwo .predavatelji .predavatelj .naziv').css({
            'font-size' : '11px',
            'height' : '40px',
            'padding' : '5px 10px',
            'font-weight' : '400'
          });
          if($('.event-info #ifmReport').contents().find('#SessionTwo #predavateljiAll').html() == ""){
            $('.event-info #ifmReport').contents().find('#SessionTwo .predavanje').css({
              'display': 'block',
              'width': 'calc(100% - 30px)',
              'margin-left' : '30px'
            });
          }

          //once the event content has been loaded
          self.element.addClass('content-loaded');
        });
        this.modalBody.find('.event-info #ifmReport').attr('src', event.parent().attr('data-content'));

        this.element.addClass('modal-is-open');

        setTimeout(() => {
          //fixes a flash when an event is selected - desktop version only
          event.parent('li').addClass('selected-event');
        }, 10);

        if( mq == 'mobile' ) {
          self.modal.one(this.transitionEnd, function(){
            self.modal.off(this.transitionEnd);
            self.animating = false;
          });
        } else {
          var eventTop = event.offset().top - $(window).scrollTop(),
            eventLeft = event.offset().left,
            eventHeight = event.innerHeight(),
            // eventWidth = event.innerWidth();
            eventWidth = 150;

          var windowWidth = $(window).width(),
            windowHeight = $(window).height();

          var modalWidth = ( windowWidth*.8 > self.modalMaxWidth ) ? self.modalMaxWidth : windowWidth*.8,
            modalHeight = ( windowHeight*.8 > self.modalMaxHeight ) ? self.modalMaxHeight : windowHeight*.8;

          var modalTranslateX = (windowWidth - modalWidth)/2 - eventLeft,
            modalTranslateY = (windowHeight - modalHeight)/2 - eventTop;

          var HeaderBgScaleY = modalHeight/eventHeight,
            BodyBgScaleX = (modalWidth - eventWidth);

          //change modal height/width and translate it
          self.modal.css({
            top: eventTop+'px',
            left: eventLeft+'px',
            height: modalHeight+'px',
            width: modalWidth+'px',
          });
          this.transformElement(self.modal, 'translateY('+modalTranslateY+'px) translateX('+modalTranslateX+'px)');

          //set modalHeader width
          self.modalHeader.css({
            width: eventWidth+'px',
          });
          //set modalBody left margin
          self.modalBody.css({
            marginLeft: eventWidth+'px',
          });

          //change modalBodyBg height/width ans scale it
          self.modalBodyBg.css({
            height: eventHeight+'px',
            width: '1px',
          });
          this.transformElement(self.modalBodyBg, 'scaleY('+HeaderBgScaleY+') scaleX('+BodyBgScaleX+')');

          //change modal modalHeaderBg height/width and scale it
          self.modalHeaderBg.css({
            height: eventHeight+'px',
            width: eventWidth+'px',
          });
          this.transformElement(self.modalHeaderBg, 'scaleY('+HeaderBgScaleY+')');

          self.modalHeaderBg.one(this.transitionEnd, function(){
            //wait for the  end of the modalHeaderBg transformation and show the modal content
            self.modalHeaderBg.off(this.transitionEnd);
            self.animating = false;
            self.element.addClass('animation-completed');
          });
        }

        //if browser do not support transitions -> no need to wait for the end of it
        if( !this.transitionsSupported ) self.modal.add(self.modalHeaderBg).trigger(this.transitionEnd);
    }

    private closeModal(event) {
        var self = this;
        var mq = self.mq();

        this.animating = true;

        if( mq == 'mobile' ) {
          this.element.removeClass('modal-is-open');
          this.modal.one(this.transitionEnd, function(){
            self.modal.off(this.transitionEnd);
            self.animating = false;
            self.element.removeClass('content-loaded');
            event.removeClass('selected-event');
          });
        } else {
          var eventTop = event.offset().top - $(window).scrollTop(),
            eventLeft = event.offset().left,
            eventHeight = event.innerHeight(),
            eventWidth = event.innerWidth();

          var modalTop = Number(self.modal.css('top').replace('px', '')),
            modalLeft = Number(self.modal.css('left').replace('px', ''));

          var modalTranslateX = eventLeft - modalLeft,
            modalTranslateY = eventTop - modalTop;

          self.element.removeClass('animation-completed modal-is-open');

          //change modal width/height and translate it
          this.modal.css({
            width: eventWidth+'px',
            height: eventHeight+'px'
          });
          this.transformElement(self.modal, 'translateX('+modalTranslateX+'px) translateY('+modalTranslateY+'px)');

          //scale down modalBodyBg element
          this.transformElement(self.modalBodyBg, 'scaleX(0) scaleY(1)');
          //scale down modalHeaderBg element
          this.transformElement(self.modalHeaderBg, 'scaleY(1)');

          this.modalHeaderBg.one(this.transitionEnd, function(){
            //wait for the  end of the modalHeaderBg transformation and reset modal style
            self.modalHeaderBg.off(this.transitionEnd);
            self.modal.addClass('no-transition');
            setTimeout(() => {
              self.modal.add(self.modalHeader).add(self.modalBody).add(self.modalHeaderBg).add(self.modalBodyBg).attr('style', '');
            }, 10);
            setTimeout(() => {
              self.modal.removeClass('no-transition');
            }, 20);

            self.animating = false;
            self.element.removeClass('content-loaded');
            event.removeClass('selected-event');
          });
        }

        //browser do not support transitions -> no need to wait for the end of it
        if(!this.transitionsSupported) self.modal.add(self.modalHeaderBg).trigger(this.transitionEnd);
    }

    private mq() {
        //get MQ value ('desktop' or 'mobile')
        var self = this;
        return window.getComputedStyle(this.element.get(0), '::before').getPropertyValue('content').replace(/["']/g, '');
    }

    private checkEventModal() {
        this.animating = true;
        var self = this;
        var mq = this.mq();

        if( mq == 'mobile' ) {
            //reset modal style on mobile
            self.modal.add(self.modalHeader).add(self.modalHeaderBg).add(self.modalBody).add(self.modalBodyBg).attr('style', '');
            self.modal.removeClass('no-transition');
            self.animating = false;
        } else if( mq == 'desktop' && self.element.hasClass('modal-is-open') ) {
            self.modal.addClass('no-transition');
            self.element.addClass('animation-completed');
            var event = self.eventsGroup.find('.selected-event');

            var eventTop = event.offset().top - $(window).scrollTop(),
            eventLeft = event.offset().left,
            eventHeight = event.innerHeight(),
            // eventWidth = event.innerWidth();
            eventWidth = 150;

            var windowWidth = $(window).width(),
            windowHeight = $(window).height();

            var modalWidth = ( windowWidth*.8 > self.modalMaxWidth ) ? self.modalMaxWidth : windowWidth*.8,
            modalHeight = ( windowHeight*.8 > self.modalMaxHeight ) ? self.modalMaxHeight : windowHeight*.8;

            var HeaderBgScaleY = modalHeight/eventHeight,
            BodyBgScaleX = (modalWidth - eventWidth);

            setTimeout(function(){
            self.modal.css({
                width: modalWidth+'px',
                height: modalHeight+'px',
                top: (windowHeight/2 - modalHeight/2)+'px',
                left: (windowWidth/2 - modalWidth/2)+'px',
            });
            this.transformElement(self.modal, 'translateY(0) translateX(0)');
            //change modal modalBodyBg height/width
            self.modalBodyBg.css({
                height: modalHeight+'px',
                width: '1px',
            });
            this.transformElement(self.modalBodyBg, 'scaleX('+BodyBgScaleX+')');
            //set modalHeader width
            self.modalHeader.css({
                width: eventWidth+'px',
            });
            //set modalBody left margin
            self.modalBody.css({
                marginLeft: eventWidth+'px',
            });
            //change modal modalHeaderBg height/width and scale it
            self.modalHeaderBg.css({
                height: eventHeight+'px',
                width: eventWidth+'px',
            });
            this.transformElement(self.modalHeaderBg, 'scaleY('+HeaderBgScaleY+')');
            }, 10);

            setTimeout(() => {
            self.modal.removeClass('no-transition');
            self.animating = false;
            }, 20);
        }
    }

    private getScheduleTimestamp(time) {
        //accepts hh:mm format - convert hh:mm to timestamp
        time = time.replace(/ /g,'');
        var timeArray = time.split(':');
        var timeStamp = parseInt(timeArray[0])*60 + parseInt(timeArray[1]);
        return timeStamp;
    }

    private transformElement(element, value) {
        element.css({
            '-moz-transform': value,
            '-webkit-transform': value,
          '-ms-transform': value,
          '-o-transform': value,
          'transform': value
        });
    }
}
