/**
 * @description подписка юзеров на события календаря
 * @author KirilenkoAS
 * @date 18.04.2019
 * @version 1.0.0
 */

SP.SOD.executeOrDelayUntilScriptLoaded(loaded, "sp.ribbon.js");

function loaded() {

	$(document).ready(function () {
		var selected = SP.ListOperation.Selection.getSelectedItems(ctx);
		console.log(selected == SP.ListOperation.Selection.getSelectedItems);

		/**
		 * @constant
		 */
		var SUBSCRIBERS_LIST = 'FBDD788C-D8A1-456C-B019-BF22977E095E', // ID списка подписчиков на календарь
			CALENDAR_LIST = 'DE850819-F0D3-4060-81E7-57C4838BAA32', // ID списка календаря
			CALENDAR_NAME = 'Calendar',
			ROOT_URL = ctx.HttpRoot,
			CURRENT_USER_ID = ctx.CurrentUserId;
		// фразы
		var TEXT_ERROR = 'Произошла ошибка. Обновите страницу и повторите попытку.',
			TEXT_SELECT_EVENT = 'Выберите одно событие.',
			TEXT_NETWORK_ERROR = 'Возникла сетевая ошибка. Обновите страницу и попробуйте снова.',
			TEXT_CALENDAR_UNSUBSCRIBE = 'Вы отписались от рассылки календаря.',
			TEXT_CALENDAR_SUBSCRIBE = 'Вы подписались на рассылку календаря.',
			TEXT_EVENT_UNSUBSCRIBE = 'Вы отписались от этого события.',
			TEXT_EVENT_SUBSCRIBE = 'Вы подписались на это событие.';
	
		function Subscriber() {
			this.headers = {
				'Content-Type': 'application/json; odata=verbose',
				'Accept': 'application/json; odata=verbose'
			},

			this.eventsID = null;
			this.isCalendarSubscriber = null;
			this.isEventsSubscriber = null;
			//this.selectedItems = null;
			//this.selectedItemID = null;
			//this.currentUser = ctx.CurrentUserId;

		}

		Subscriber.prototype.init = function() {
			// доступность вкладки для инициализации рибона, хз че там происходит, но ведет он себя крайне странно
			var ribbonInitTab = _ribbon.initialTabId;
			// проверяем доступность объекта риббона
			if ( ribbonInitTab !== null ) {
				// вставляем блок прелоадера
				this.insertPreloaderDOM();
				// добавляем вкладки
				this.subscribeTabInit();
				// регистрируем события на кнопки
				this.registerButtonEvents();
				// асинхронные операции
				this.asyncInitData();
			}		
		}

		Subscriber.prototype.insertPreloaderDOM = function() {
			$('body').append('<div id="SubscribersPreloader"><img src="' + ROOT_URL + '/SiteAssets/img/45.gif" /></div>');
		}

		Subscriber.prototype.createSbuscribeTab = function() {
			var ribbon = SP.Ribbon.PageManager.get_instance().get_ribbon();
			if (ribbon !== null) {
				// данные по кастомным вкладкам
				var tabs = [
					{
						id: 'Subscribe',
						title: 'Подписка',
						desc: 'Подписка на события и календарь',
						groups: [
							// подписка на события
							{ id: 'SubscribeEvent', title: 'Подписка на событие', desc: 'Нажмите, чтобы получать уведомления по выбранному событию' },
							// подписка на весь календарь
							{ id: 'SubscribeAll', title: 'Подписка на календарь', desc: 'Нажмите, чтобы получать уведомления по каждому событию календаря' }
						]
					}
				];
				// регистрация табов
				for (var i = 0; i < tabs.length; i++) {
					var element = tabs[i];
					tab = new CUI.Tab( ribbon, element.id + '.Tab', element.title, element.desc, element.id + '.Tab.Command', false, '', null );
					ribbon.addChildAtIndex(tab, i + 1);
					// регистрация кнопок
					for (var i = 0; i < element.groups.length; i++) {
						var group = element.groups[i];
						var groupInsert = new CUI.Group( ribbon, group.id + '.Tab.Group', group.title, group.desc, group.id + '.Group.Command', null );
						tab.addChild(groupInsert);
					}
				}
			}
			SelectRibbonTab('Subscribe.Tab', true);

			var html = '<a href="#" id="SubscribeEvent" ><img src="/kirilenkoas/SiteAssets/img/notify.png" /></a>';
			$('#SubscribeEvent\\.Tab\\.Group .ms-cui-groupBody').html( html );

			var html = '<a href="#" id="SubscribeAll" ><img src="/kirilenkoas/SiteAssets/img/notify-1.png" /></a>';
			$('#SubscribeAll\\.Tab\\.Group .ms-cui-groupBody').html( html );
		}

		Subscriber.prototype.subscribeTabInit = function() {
			var _self = this;
			var pm = SP.Ribbon.PageManager.get_instance();

			/**
			 * @description событие, выполняется при инициализации вкладки Ribbon.Read
			 */
			pm.add_ribbonInited(function () {
				_self.createSbuscribeTab();
			});

			var ribbon = null;
			try {
				ribbon = pm.get_ribbon();
			}
			catch ( e ) { }
			
			if ( !ribbon ) {
				if ( typeof ( _ribbonStartInit ) == "function" ) {
					_ribbonStartInit( _ribbon.initialTabId, false, null );
				}
			}
			else {
				this.createSbuscribeTab();
			}
		}

		// Subscriber.prototype.ribbonObserver = function() {
		// 	$('#RibbonContainer').on("DOMNodeInserted", function (event) {
		// 		var subWrap = $(event.target).find('#Ribbon\\.Calendar\\.Events\\.Share-LargeMedium-0-0')
		// 		if ( $(subWrap).length ) {
		// 			if ( $('#SubscribeToEvent').length === 0 ) {
		// 				$(subWrap).append('<a id="SubscribeToEvent" href="javascript:;">Подписка</a>');
		// 			}
		// 		}
		// 	});
		// }

		Subscriber.prototype.asyncError = function( error ) {
			console.log( error );
			alert( TEXT_ERROR )
		}

		Subscriber.prototype.hidePreloader = function() {
			$('#SubscribersPreloader').hide(0);
			console.log('Асинхронная операция завершена...');
		}

		Subscriber.prototype.showPreloader = function() {
			$('#SubscribersPreloader').show(0);
			console.log('Асинхронная операция запущена...');
		}

		Subscriber.prototype.asyncInitData = function() {
			var _self = this;

			$.when( this.getCalendarSubscriber(), this.getEventsIDsByUser() )
				.then( function() {
					_self.setCalendarButtonClass();
				})
				.fail( this.asyncError )
				.always( this.hidePreloader );
		}

		/**
		 * @description добавляет классы кнопке подписки на календарь
		 */
		Subscriber.prototype.setCalendarButtonClass = function() {
			if ( this.isCalendarSubscriber.status ) {
				$('#SubscribeAll\\.Tab\\.Group .ms-cui-groupBody').removeClass( 'unsubscribed' ).addClass( 'subscribed' );
			} else {
				$('#SubscribeAll\\.Tab\\.Group .ms-cui-groupBody').removeClass( 'subscribed' ).addClass( 'unsubscribed' );
			}
		}

		/**
		 * @description добавление классов кнопке подписки на событие
		 */
		Subscriber.prototype.setEventButtonClass = function() {
			var selectedItems = SP.ListOperation.Selection.getSelectedItems(ctx);
			var itemID = selectedItems[0].id;
			// если ID не является числом
			if ( isNaN( +itemID ) ) {
				var itemID = itemID.split( '.' )[0];
			}
			if ( itemID in this.eventsID ) {
				var index = this.eventsID[itemID].indexOf( +CURRENT_USER_ID );
				if ( index !== -1 ) {
					$('#SubscribeEvent\\.Tab\\.Group .ms-cui-groupBody').removeClass( 'unsubscribed' ).addClass( 'subscribed' );
				} else {
					$('#SubscribeEvent\\.Tab\\.Group .ms-cui-groupBody').removeClass( 'subscribed' ).addClass( 'unsubscribed' );
				}
			} else {
				$('#SubscribeEvent\\.Tab\\.Group .ms-cui-groupBody').removeClass( 'subscribed' ).addClass( 'unsubscribed' );
			}
		}

		/**
		 * @todo Проработаь косяк, если элемент удаляется вне страницы
		 */
		Subscriber.prototype.registerButtonEvents = function() {
			var _self = this;
			
			// при клике на событие календаря
			$('body').on('click', '.ms-acal-rootdiv .ms-acal-item', function() {
				setTimeout(function() {
					SelectRibbonTab('Subscribe.Tab', true);
					_self.setEventButtonClass();
				}, 100);
			});

			// при клике в любом месте, кроме события календаря
			$(document).click(function (e){
				var item = $(".ms-acal-rootdiv .ms-acal-item");
				var ribbon = $('#s4-ribbonrow');
				if ( !item.is( e.target ) && item.has( e.target ).length === 0
							&& !ribbon.is( e.target ) && ribbon.has( e.target ).length === 0 ) {
					$('#SubscribeEvent\\.Tab\\.Group .ms-cui-groupBody').removeClass( 'unsubscribed' ).removeClass( 'subscribed' );
				}
			});

			// подписаться/отписаться на событие
			$('body').on('click', '#SubscribeEvent', function() {
				//var currentUser = ctx.CurrentUserId;
				var selectedItems = SP.ListOperation.Selection.getSelectedItems(ctx);
				
				if ( selectedItems.length == 1 ) {
					var selectedItemID = selectedItems[0].id;
					_self.toggleActionToEvent( selectedItemID );
				} else {
					alert( TEXT_SELECT_EVENT );
				}
			});

			// подписаться/отписаться на календарь
			$('body').on('click', '#SubscribeAll', function() {
				if ( _self.isCalendarSubscriber !== null ) {
					_self.showPreloader();
					// если юзер подписан на календарь
					if ( _self.isCalendarSubscriber.status ) {
						// удаляем элемент в списке - отписываемся
						_self.unsubscribeFromCalendar()
					} else {
						// добавляем элемент в список - подписываемся
						_self.subscribeToCalendar()
					}
				} else {
					alert( TEXT_NETWORK_ERROR );
				}
			});
		}

		/**
		 * @description получение юзера из списка подписчиков календаря
		 */
		Subscriber.prototype.getCalendarSubscriber = function() {
			var url = ROOT_URL + "/_api/web/lists/GetById('" + SUBSCRIBERS_LIST + "')/items?$filter=(fldSubscriber/Id eq '" + CURRENT_USER_ID + "' or Id eq '-1')";
			var _self = this;

			return $.ajax({ method: "GET", url: url, headers: this.headers })
						.done( function( response ) {
							if ( _self.isCalendarSubscriber === null ) _self.isCalendarSubscriber = new Object;
							try {
								var results = response.d.results;
								// если подписан
								if ( results.length ) {
									if ( _self.isCalendarSubscriber === null ) _self.isCalendarSubscriber = new Object;
									_self.isCalendarSubscriber.id = results[0].Id;
									_self.isCalendarSubscriber.status = true;
								} else {
									_self.isCalendarSubscriber.id = null;
									_self.isCalendarSubscriber.status = false;
								}
							} catch (error) {
								_self.asyncError( error );
							}
						})
		}

		/**
		 * @description удаление юзера из списка подписчиков календаря
		 * @todo возможно, зарефактрить unsubscribeFromCalendar и subscribeToCalendar
		 */
		Subscriber.prototype.unsubscribeFromCalendar = function() {
			var _self = this;
			var url = ROOT_URL + "/_api/web/lists/GetById('" + SUBSCRIBERS_LIST + "')/items(" + this.isCalendarSubscriber.id + ")";
			var _headers = {
				'Content-Type': 'application/json; odata=verbose',
				'Accept': 'application/json; odata=verbose',
				'X-HTTP-Method': 'DELETE',
				'If-Match': '*'
			}
			return _self.getRefreshDigest()
						.then( function( digest ) {
							_headers['X-RequestDigest'] = digest.d.GetContextWebInformation.FormDigestValue;
							return $.ajax({ method: 'POST', url: url, headers: _headers })
								.done( function( r, a, x ) {
									try {
										if ( x.status === 200 ) {
											_self.isCalendarSubscriber.id = null;
											_self.isCalendarSubscriber.status = false;
											_self.setCalendarButtonClass();
											alert( TEXT_CALENDAR_UNSUBSCRIBE )
										}
									} catch (error) {
										this.asyncError( error );
									}
								})
						})
						.fail( _self.asyncError )
						.always( _self.hidePreloader );
		}

		/**
		 * @description добавление элемента в списко подписчиков календаря
		 * @todo возможно, зарефактрить unsubscribeFromCalendar и subscribeToCalendar
		 */
		Subscriber.prototype.subscribeToCalendar = function( digest ) {
			var _self = this;
			var url = ROOT_URL + "/_api/web/lists/GetById('" + SUBSCRIBERS_LIST + "')/items";
			var _headers = {
				'Content-Type': 'application/json; odata=verbose',
				'Accept': 'application/json; odata=verbose',
			}
			var _data = JSON.stringify( { __metadata: { type: 'SP.ListItem' }, fldSubscriberId: CURRENT_USER_ID } )
			return _self.getRefreshDigest()
						.then( function( digest ) {
							_headers['X-RequestDigest'] = digest.d.GetContextWebInformation.FormDigestValue;
							return $.ajax({ method: 'POST', url: url, headers: _headers, data: _data })
								.done( function( r, a, x ) {
									try {
										if ( 'Id' in r.d ) {
											_self.isCalendarSubscriber.id = r.d.Id;
											_self.isCalendarSubscriber.status = true;
											_self.setCalendarButtonClass();
											alert( TEXT_CALENDAR_SUBSCRIBE )
										}
									} catch (error) {
										this.asyncError( error );
									}
								})
						})
						.fail( _self.asyncError )
						.always( _self.hidePreloader );
		}

		/**
		 * @description получение ID событий, на которые подписан юзер... с помощью CAML Query
		 * @see https://docs.microsoft.com/ru-ru/sharepoint/dev/schema/collaborative-application-markup-language-caml-schemas
		 */
		Subscriber.prototype.getEventsIDsByUser = function() {
			var _self = this;
			var dateNowISO = new Date().toISOString();
			var filter = "fldSubscribers/Id eq '" + CURRENT_USER_ID + "' and fldEventStart ge datetime'" + dateNowISO + "' or Id eq '-1'";
			var url = ROOT_URL + "/_api/web/lists/GetById('" + CALENDAR_LIST + "')/getitems";
			// var viewXml = '<View><Query><Where><And><Eq><FieldRef Name="fldSubscribers" LookupId="TRUE" /><Value Type="Lookup">' + CURRENT_USER_ID + '</Value></Eq><Geq><FieldRef Name="EventDate" /><Value Type="DateTime" IncludeTimeValue="TRUE">' + dateNowISO + '</Value></Geq></And></Where></Query></View>';
			var viewXml = '<View><Query><Where><Eq><FieldRef Name="fldSubscribers" LookupId="TRUE" /><Value Type="Lookup">' + CURRENT_USER_ID + '</Value></Eq></Where></Query></View>';
			var queryPayload = JSON.stringify({  
				'query' : {
					   '__metadata': { 'type': 'SP.CamlQuery' }, 
					   'ViewXml' : viewXml  
				}
			 });
			 var _headers = {
				'Content-Type': 'application/json; odata=verbose',
				'Accept': 'application/json; odata=verbose',
				'X-RequestDigest': $("#__REQUESTDIGEST").val()
			}
			return $.ajax({ method: "POST", url: url, headers: _headers, data: queryPayload })
				.done( function( response ) {
					try {
						_self.eventsID = new Object;
						for (var i = 0; i < response.d.results.length; i++) {
							var event = response.d.results[i];
							_self.eventsID[event.Id] = event.fldSubscribersId.results;
						}
					} catch (error) {
						_self.asyncError( error );
					}
				})
		}

		/**
		 * @description получение юзеров из события по ID
		 * @param {Number} itemID
		 * @returns {Promise}
		 */
		Subscriber.prototype.getEventSubscribersByID = function( itemID ) {
			var url = ROOT_URL + "/_api/web/lists/GetById('" + CALENDAR_LIST + "')/items(" + itemID + ")";
			return $.ajax({ method: "GET", url: url, headers: this.headers })
						.then( function( res ) {
							if ( res.d.fldSubscribersId ) {
								return res.d.fldSubscribersId.results
							} else {
								return new Array;
							}
						})
						.fail( this.asyncError );
		}

		/**
		 * @description отписка юзера от выбранного события
		 * @param {Number} itemID выбранный элемент в списке
		 * @returns {Promise}
		 */
		Subscriber.prototype.toggleActionToEvent = function( itemID ) {
			// ! из-за косяка SharePoint`а в виде выбранного itemID может прийти "7.0.2019-04-30T11:00:00Z"
			// если ID не является числом
			if ( isNaN( +itemID ) ) {
				var itemID = itemID.split( '.' )[0];
			}
			// получаем актуальные данные по подписчикам события
			return this.getEventsIDsByUser()
						.then( this.getRefreshDigest )
						.then( this.calculatedEventSubscribers.bind( this, itemID ) )
						.fail( this.asyncError )
						.always( this.hidePreloader );
		}

		/**
		 * @description основная логика определения подписки\отписки юзера от события
		 * @param {Number} itemID
		 * @param {Object} digest
		 */
		Subscriber.prototype.calculatedEventSubscribers = function( itemID, digest ) {
			var _self = this;
			var ajaxParams = {
				url: ROOT_URL + "/_api/web/lists/GetById('" + CALENDAR_LIST + "')/items(" + itemID + ")",
				headers: {
					'Content-Type': 'application/json; odata=verbose',
					'Accept': 'application/json; odata=verbose',
					'X-HTTP-Method': 'MERGE',
					'X-RequestDigest': digest.d.GetContextWebInformation.FormDigestValue,
					'If-Match': '*'
				}
			}
			var tempSubscribers;
			// если юзер подписан
			if ( itemID in _self.eventsID ) {
				var index = _self.eventsID[itemID].indexOf( +CURRENT_USER_ID );
				// если юзер все еще подписан на выбранное событие
				if ( index !== -1 ) {
					// копируем состояние массива событий, на которые подписан юзер
					tempSubscribers = [].concat( _self.eventsID[itemID] );
					tempSubscribers.splice( index, 1 );
					ajaxParams.users = tempSubscribers;
					// коллбэк после обновления
					var sendMessage = function() {
						alert( TEXT_EVENT_UNSUBSCRIBE );
					}
					// обновляем данные на сервере
					_self.updateEventSubscribers( itemID, ajaxParams, sendMessage );
				} else {
					alert( TEXT_EVENT_UNSUBSCRIBE );
					return false;
				}
			// если юзер не подписан
			} else {
				_self.getEventSubscribersByID( itemID )
					.done( function( subscribers ) {
						if ( subscribers instanceof Array ) {
							// добавляем юзера в массив подписчиков
							tempSubscribers = subscribers;
							tempSubscribers.push( CURRENT_USER_ID );
							ajaxParams.users = tempSubscribers;
							// коллбэк после обновления
							var sendMessage = function() {
								alert( TEXT_EVENT_SUBSCRIBE );
							}
							// обновляем данные на сервере
							_self.updateEventSubscribers( itemID, ajaxParams, sendMessage );
						}
					})
			}
		}

		/**
		 * @description обновление события с новой инфой о подписчиках
		 * @param {Number} itemID
		 * @param {Object} params объект с параметрами запроса: url, headers, users (id)
		 * @param {Function} callback коллбек при обновление данных
		 */
		Subscriber.prototype.updateEventSubscribers = function( itemID, params, callback ) {
			var _self = this;
			var _data = JSON.stringify({
				__metadata: { type: 'SP.Data.' + CALENDAR_NAME + 'ListItem' },
				fldSubscribersId: { results: params.users }
			});

			$.ajax({ method: 'POST', url: params.url, headers: params.headers, data: _data })
				.done( function( r, a, x ) {
					try {
						if ( x.status === 204 ) {
							_self.eventsID[itemID] = params.users;
							_self.setEventButtonClass();
							callback();
						}
					} catch (error) {
						_self.asyncError( error );
					}
				})
		}

		/**
         *
         * Загружает с сервера X-RequestDigest
         *
         * @returns {Promise}
         */
        Subscriber.prototype.getRefreshDigest = function() {
			var url = ROOT_URL + '/_api/contextinfo';
			return $.ajax({ method: "POST", url: url, headers: this.headers });
        }

		var subscriber = new Subscriber;
		subscriber.init();

	});
}