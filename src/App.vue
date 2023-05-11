<template>
  <v-app>
    <v-main>
      <template>
        <v-container fluid>
          <v-form>
            <v-col
                cols="12"
                sm="6"
                md="4"
            >
              <v-menu
                  ref="menu"
                  v-model="menu"
                  :close-on-content-click="false"
                  transition="scale-transition"
                  offset-y
                  min-width="auto"
              >
                <template v-slot:activator="{ on, attrs }">
                  <v-text-field
                      v-model="dates"
                      label="Дата"
                      prepend-icon="mdi-calendar"
                      readonly
                      v-bind="attrs"
                      v-on="on"
                      outlined
                      dense
                  ></v-text-field>
                </template>
                <v-date-picker
                    v-model="dates"
                    range
                    locale="ru-RU"
                    no-title
                >
                  <v-spacer></v-spacer>
                  <v-btn
                      text
                      color="primary"
                      @click="menu = false"
                  >
                    Отмена
                  </v-btn>
                  <v-btn
                      text
                      color="primary"
                      @click="$refs.menu.save(dates)"
                  >
                    Сохранить
                  </v-btn>
                </v-date-picker>
              </v-menu>
            </v-col>
            <v-col class="ma-2">
              <tempate v-if="isMobile">
                <v-row>
                  <v-col>
                    <v-btn
                        class="pa-10"
                        @click="getReport"
                        :disabled="loading"
                        elevation="1"
                    >
                      Сформировать
                    </v-btn>
                  </v-col>
                  <v-col>
                    <v-btn
                        elevation="1"
                        class="pa-10"
                        @click="exportXls"
                        :disabled="isXLS"
                    >
                      Выгрузить XLS
                    </v-btn>
                  </v-col>
                </v-row>
              </tempate>
              <template v-else>
                <v-row>
                  <v-col cols="2" class="mr-4">
                    <v-btn
                        class="pa-10"
                        @click="getReport"
                        :disabled="loading"
                        elevation="1"
                    >
                      Сформировать
                    </v-btn>
                  </v-col>
                  <v-col cols="2">
                    <v-btn
                        elevation="1"
                        class="pa-10"
                        @click="exportXls"
                        :disabled="isXLS"
                    >
                      Выгрузить XLS
                    </v-btn>
                  </v-col>
                </v-row>
              </template>
            </v-col>
          </v-form>
          <v-row v-if="loading">
            <v-col class="text-center">
              <v-progress-linear
                  color="primary"
                  v-model="progressValue"
                  :buffer-value="bufferValue"
              ></v-progress-linear>
            </v-col>
          </v-row>
          <v-row v-if="load">
            <v-col>
              <h2>Менеджеры</h2>
              <adamantTable :items="itemsManager" :deals="totalManager"/>
            </v-col>
          </v-row>
          <v-divider class="divider" v-if="load"></v-divider>
          <v-row v-if="load">
            <v-col>
              <h2>Промоутеры</h2>
              <adamantTable :items="itemsPromo" :deals="totalPromo"/>
            </v-col>
          </v-row>
          <v-divider class="divider" v-if="load"></v-divider>
          <v-row v-if="load">
            <v-col>
              <h2>Источники</h2>
              <adamantTable :items="itemsSources" :deals="totalSourse" :menu.sync="butsMenu" :promoMenu.sync="promoMenu"
                            :otherMenu.sync="otherMenu"/>
            </v-col>
          </v-row>
          <v-divider class="divider" v-if="load"></v-divider>
        </v-container>
      </template>
    </v-main>
  </v-app>
</template>
<script>

import adamantTable from "@/components/adamantTable";
import axios from 'axios'
import exceljs from 'exceljs'
import fileWrite from 'file-saver'

export default {
  name: 'App',
  components: {adamantTable},
  data: () => ({
    progressValue: 10,
    bufferValue: 20,
    load: false,
    isMobile: false,
    isXLS: true,
    butsMenu: false,
    otherMenu: false,
    promoMenu: false,
    dates: [],
    normalDate: [],
    loading: true,
    menu: false,
    itemsManager: [],
    itemsPromo: [],
    itemsSources: [],
    managers1: [],
    managers2: [],
    sourse: [],
    buts: [],
    promoManagers: [],
    totalManager: {
      win: 0,
      work: 0,
      close: 0,
      tours: 0,
      doubleTours: 0,
      toursPlan: 0,
      totalDeals: 0,
      color: 'orange lighten-3'
    },
    totalPromo: {
      win: 0,
      work: 0,
      close: 0,
      tours: 0,
      doubleTours: 0,
      toursPlan: 0,
      totalDeals: 0,
      color: 'orange lighten-3'
    },
    totalSourse: {
      win: 0,
      work: 0,
      close: 0,
      tours: 0,
      doubleTours: 0,
      toursPlan: 0,
      totalDeals: 0,
      color: 'orange lighten-3'
    },
    deals: {
      win: [],
      work: [],
      close: [],
      tours: [],
      doubleTours: [],
      toursPlan: [],
      totalDeals: []
    },
    dealsToSource: {
      win: [],
      work: [],
      close: [],
      tours: [],
      doubleTours: [],
      toursPlan: [],
      totalDeals: []
    },
    monthMatrix: {
      1: 31,
      2: 28,
      3: 31,
      4: 30,
      5: 31,
      6: 30,
      7: 31,
      8: 31,
      9: 30,
      10: 31,
      11: 30,
      12: 31,
    },
    filterStageWork: [
      'NEW', 'UC_B9M183', 'UC_IIQTOU', 'UC_O34O23', 'PREPARATION', 'UC_L76NCG', 'UC_6NLBBE', 'UC_O5W0X2', 'UC_ZABP36', 'EXECUTING'
    ],
    filterStageClose: [
      '6', '5', '4', '3', '2', '1', 'APOLOGY', 'LOSE', 'C4:LOSE'
    ],
    stageWin: ['WON', 'C4:WON'],
    stageToursPlan: ['UC_6NLBBE', 'C4:UC_8RZB1U'],
    url: 'https://andamanriviera.bitrix24.ru/rest/224/txcul21bs9jk2n34/'
  }),
  methods: {
    async normalizeDate() {
      try {
        console.log(this.dates, 'dates')
        const firstDate = this.dates[0].split('-')
        const twoDate = this.dates[1].split('-')
        const firstDateFormated = firstDate[2] + '.' + firstDate[1] + '.' + firstDate[0] + ' 00:00:00'
        const twoDateFormated = twoDate[2] + '.' + twoDate[1] + '.' + twoDate[0] + ' 23:59:59'
        this.normalDate.push(firstDateFormated)
        this.normalDate.push(twoDateFormated)
      } catch (e) {
        console.log(e)
      }
    },
    async pause(ms) {
      return new Promise((resolve) => {
        setTimeout(resolve, ms);
      });
    },
    async getReport() {
      try {
        this.totalManager = {
          win: 0,
          work: 0,
          close: 0,
          tours: 0,
          doubleTours: 0,
          toursPlan: 0,
          totalDeals: 0,
          color: 'orange lighten-3'
        }
        this.totalPromo = {
          win: 0,
          work: 0,
          close: 0,
          tours: 0,
          doubleTours: 0,
          toursPlan: 0,
          totalDeals: 0,
          color: 'orange lighten-3'
        }
        this.totalSourse = {
          win: 0,
          work: 0,
          close: 0,
          tours: 0,
          doubleTours: 0,
          toursPlan: 0,
          totalDeals: 0,
          color: 'orange lighten-3'
        }
        this.deals = {
          win: [],
          work: [],
          close: [],
          tours: [],
          doubleTours: [],
          toursPlan: [],
          totalDeals: []
        }
        this.dealsToSource = {
          win: [],
          work: [],
          close: [],
          tours: [],
          doubleTours: [],
          toursPlan: [],
          totalDeals: []
        }
        this.normalDate = []
        this.itemsManager = []
        this.itemsPromo = []
        this.itemsSources = []
        this.progressValue = 5
        this.bufferValue = 5
        this.loading = true
        this.load = false
        this.normalizeDate()
        //получаем сделки для отчета по менеджерам
        const dealsWork = await this.getDealsFunc('crm.stagehistory.list', {
          entityTypeId: 2,
          filter: {
            '>=CREATED_TIME': this.normalDate[0],
            '<=CREATED_TIME': this.normalDate[1],
            CATEGORY_ID: 0,
            TYPE_ID: [1, 2]
          },
          select: ['OWNER_ID'],
          start: 0
        })
        this.progressValue += 5
        this.bufferValue += 5
        const dealsClose = await this.getDealsFunc('crm.stagehistory.list', {
          entityTypeId: 2,
          filter: {
            '>=CREATED_TIME': this.normalDate[0],
            '<=CREATED_TIME': this.normalDate[1],
            CATEGORY_ID: 0,
            TYPE_ID: 3,
            STAGE_ID: this.filterStageClose
          },
          select: ['OWNER_ID'],
          start: 0
        }, dealsWork)
        this.progressValue += 5
        this.bufferValue += 5
        const dealsWin = await this.getDealsFunc('crm.stagehistory.list', {
          entityTypeId: 2,
          filter: {
            '>=CREATED_TIME': this.normalDate[0],
            '<=CREATED_TIME': this.normalDate[1],
            CATEGORY_ID: 0,
            TYPE_ID: 3,
            STAGE_ID: this.stageWin
          },
          select: ['OWNER_ID'],
          start: 0
        }, dealsWork)
        this.progressValue += 5
        this.bufferValue += 5
        const dealsPlanTour = await this.getDealsListFunc('crm.deal.list', {
          order: {},
          filter: {
            '>=UF_CRM_1671953015183': this.normalDate[0],
            '<=UF_CRM_1671953015183': this.normalDate[1],
            STAGE_ID: 'UC_6NLBBE',
            CATEGORY_ID: 0,
            CLOSED: 'N'
          },
          select: ["ID", "STAGE_ID", "SOURCE_ID", "UF_CRM_63944FE4EC0CD", "UF_CRM_639FE7BFC4BBB", 'ASSIGNED_BY_ID'],
          start: 0
        })
        for (const dealPlanTour of dealsPlanTour) {
          this.deals.toursPlan.push({
            id: dealPlanTour['ID'],
            manager: dealPlanTour['ASSIGNED_BY_ID'],
            sourse: dealPlanTour['SOURCE_ID'],
            promo: dealPlanTour['UF_CRM_63944FE4EC0CD'],
            buts: dealPlanTour['UF_CRM_639FE7BFC4BBB']
          })
        }
        for (const itemPlanTour of dealsPlanTour) {
          if (!dealsWork.includes(itemPlanTour['ID'])) {
            dealsWork.push(itemPlanTour['ID'])
          }
        }
        this.progressValue += 5
        this.bufferValue += 5
        const allDealsStageWorkToEditData = await this.getDealsListFunc('crm.deal.list', {
          order: {},
          filter: {
            '>=DATE_MODIFY': this.normalDate[0],
            '<=DATE_MODIFY': this.normalDate[1],
            CATEGORY_ID: 0,
            CLOSED: 'N'
          },
          select: ['ID'],
          start: 0
        })
        for (const item of allDealsStageWorkToEditData) {
          if (!dealsWork.includes(item['ID'])) {
            dealsWork.push(item['ID'])
          }
        }
        this.progressValue += 10
        this.bufferValue += 10
        await this.bachFunc(dealsWin, 'win')
        this.progressValue += 10
        this.bufferValue += 10
        await this.bachFunc(dealsClose, 'close')
        this.progressValue += 10
        this.bufferValue += 10
        await this.bachFunc(dealsWork, 'work')
        this.progressValue += 10
        this.bufferValue += 10
        //получаем для остальных таблиц
        const dealsOther = await this.getDealsListFunc('crm.deal.list', {
          order: {"ID": "ASC"},
          filter: {
            '>=DATE_CREATE': this.normalDate[0],
            '<=DATE_CREATE': this.normalDate[1],
            CATEGORY_ID: 0,
          },
          select: ["ID", "STAGE_ID", "SOURCE_ID", "UF_CRM_63944FE4EC0CD", "UF_CRM_639FE7BFC4BBB", 'ASSIGNED_BY_ID'],
          start: 0
        })
        this.progressValue += 10
        this.bufferValue += 10
        for (const dealOther of dealsOther) {
          if (this.stageWin.includes(dealOther['STAGE_ID'])) {
            this.dealsToSource.win.push({
              id: dealOther['ID'],
              manager: dealOther['ASSIGNED_BY_ID'],
              sourse: dealOther['SOURCE_ID'],
              promo: dealOther['UF_CRM_63944FE4EC0CD'],
              buts: dealOther['UF_CRM_639FE7BFC4BBB']
            })
          } else if (this.filterStageClose.includes(dealOther['STAGE_ID'])) {
            this.dealsToSource.close.push({
              id: dealOther['ID'],
              manager: dealOther['ASSIGNED_BY_ID'],
              sourse: dealOther['SOURCE_ID'],
              promo: dealOther['UF_CRM_63944FE4EC0CD'],
              buts: dealOther['UF_CRM_639FE7BFC4BBB']
            })
          } else {
            this.dealsToSource.work.push({
              id: dealOther['ID'],
              manager: dealOther['ASSIGNED_BY_ID'],
              sourse: dealOther['SOURCE_ID'],
              promo: dealOther['UF_CRM_63944FE4EC0CD'],
              buts: dealOther['UF_CRM_639FE7BFC4BBB']
            })
          }
        }
        this.progressValue += 10
        this.bufferValue += 10
        //получаем туры из списков
        await this.getListFunc('lists.element.get', {
          IBLOCK_TYPE_ID: 'lists',
          IBLOCK_ID: 44,
          FILTER: {
            '>=PROPERTY_126': this.normalDate[0],
            '<=PROPERTY_126': this.normalDate[1],
            'PROPERTY_166': 248,
            'PROPERTY_124': 184
          }
        }, 'tours')
        await this.getListFunc('lists.element.get', {
          IBLOCK_TYPE_ID: 'lists',
          IBLOCK_ID: 44,
          FILTER: {
            '>=PROPERTY_126': this.normalDate[0],
            '<=PROPERTY_126': this.normalDate[1],
            'PROPERTY_166': 250,
            'PROPERTY_124': 184
          }
        }, 'doubleTours')

        this.getReportManagers()
        this.getReportSource()
        this.getReportPromo()
        this.getButsToSource()
        this.progressValue += 10
        this.loading = false
        this.isXLS = false
        this.load = true
      } catch (e) {
        console.log(e)
      }
    },
    async getListFunc(method, params, type) {
      try {
        const elems = []
        let isTrue = true
        while (isTrue) {
          //получаем все сделки, которые были в работе за отчетный период
          const allElemsData = await axios.post(this.url + method, params)
          await this.pause(250)
          for (const elem of allElemsData.data.result) {
            this.deals[type].push({
              id: elem['ID'],
              manager: elem['CREATED_BY'],
              sourse: (elem['PROPERTY_156']) ? Object.values(elem['PROPERTY_156'])[0] : '',
              promo: (elem['PROPERTY_160']) ? Object.values(elem['PROPERTY_160'])[0] : '',
              buts: (elem['PROPERTY_162']) ? Object.values(elem['PROPERTY_162'])[0] : ''
            })
          }
          if (allElemsData.data.next) params.start = allElemsData.data.next
          else isTrue = false
        }
        return elems
      } catch (e) {
        console.log(e)
      }
    },
    getReportManagers() {
      try {
        for (const user of this.managers1) {
          const workManagerDeals = this.deals.work.filter(e => user.ID == e.manager).length > 0 ? this.deals.work.filter(e => user.ID == e.manager).length : 0
          const closeManagerDeals = this.deals.close.filter(e => user.ID == e.manager).length > 0 ? this.deals.close.filter(e => user.ID == e.manager).length : 0
          const winManagerDeals = this.deals.win.filter(e => user.ID == e.manager).length > 0 ? this.deals.win.filter(e => user.ID == e.manager).length : 0
          const planToursManagerDeals = this.deals.toursPlan.filter(e => user.ID == e.manager).length > 0 ? this.deals.toursPlan.filter(e => user.ID == e.manager).length : 0
          const toursManagerDeals = this.deals.tours.filter(e => user.ID == e.manager).length > 0 ? this.deals.tours.filter(e => user.ID == e.manager).length : 0
          const doubleToursManagerDeals = this.deals.doubleTours.filter(e => user.ID == e.manager).length > 0 ? this.deals.doubleTours.filter(e => user.ID == e.manager).length : 0
          const dealsAllManagerDeals = Number(workManagerDeals) + Number(closeManagerDeals) + Number(winManagerDeals)

          const manager = {
            manager: `${user.LAST_NAME} ${user.NAME}`,
            dealsAll: dealsAllManagerDeals,
            dealsWork: workManagerDeals,
            dealsClose: closeManagerDeals,
            toursPlan: planToursManagerDeals,
            toursWins: toursManagerDeals,
            DoubleTours: doubleToursManagerDeals,
            dealsWins: winManagerDeals,
          }
          this.itemsManager.push(manager)

          this.totalManager.win += winManagerDeals
          this.totalManager.work += workManagerDeals
          this.totalManager.close += closeManagerDeals
          this.totalManager.tours += toursManagerDeals
          this.totalManager.doubleTours += doubleToursManagerDeals
          this.totalManager.toursPlan += planToursManagerDeals
          this.totalManager.totalDeals += dealsAllManagerDeals
        }
        for (const user of this.managers2) {
          const workManagerDeals = this.deals.work.filter(e => user.ID == e.manager).length > 0 ? this.deals.work.filter(e => user.ID == e.manager).length : 0
          const closeManagerDeals = this.deals.close.filter(e => user.ID == e.manager).length > 0 ? this.deals.close.filter(e => user.ID == e.manager).length : 0
          const winManagerDeals = this.deals.win.filter(e => user.ID == e.manager).length > 0 ? this.deals.win.filter(e => user.ID == e.manager).length : 0
          const planToursManagerDeals = this.deals.toursPlan.filter(e => user.ID == e.manager).length > 0 ? this.deals.toursPlan.filter(e => user.ID == e.manager).length : 0
          const toursManagerDeals = this.deals.tours.filter(e => user.ID == e.manager).length > 0 ? this.deals.tours.filter(e => user.ID == e.manager).length : 0
          const doubleToursManagerDeals = this.deals.doubleTours.filter(e => user.ID == e.manager).length > 0 ? this.deals.doubleTours.filter(e => user.ID == e.manager).length : 0
          const dealsAllManagerDeals = Number(workManagerDeals) + Number(closeManagerDeals) + Number(winManagerDeals)

          const manager = {
            manager: `${user.LAST_NAME} ${user.NAME}`,
            dealsAll: dealsAllManagerDeals,
            dealsWork: workManagerDeals,
            dealsClose: closeManagerDeals,
            toursPlan: planToursManagerDeals,
            toursWins: toursManagerDeals,
            DoubleTours: doubleToursManagerDeals,
            dealsWins: winManagerDeals,
          }
          this.itemsManager.push(manager)

          this.totalManager.win += winManagerDeals
          this.totalManager.work += workManagerDeals
          this.totalManager.close += closeManagerDeals
          this.totalManager.tours += toursManagerDeals
          this.totalManager.doubleTours += doubleToursManagerDeals
          this.totalManager.toursPlan += planToursManagerDeals
          this.totalManager.totalDeals += dealsAllManagerDeals
        }
        console.log(this.total, 'total')
      } catch (e) {
        console.log(e)
      }
    },
    getReportSource() {
      try {
        const other = {
          manager: 'Прочее',
          menu: false,
          dealsAll: 0,
          dealsWork: 0,
          dealsClose: 0,
          toursPlan: 0,
          toursWins: 0,
          DoubleTours: 0,
          dealsWins: 0,
          items: []
        }
        const promoSource = {
          manager: 'Реклама',
          menu: false,
          dealsAll: 0,
          dealsWork: 0,
          dealsClose: 0,
          toursPlan: 0,
          toursWins: 0,
          DoubleTours: 0,
          dealsWins: 0,
          items: []
        }
        for (const sourse of this.sourse) {
          const sourceName = sourse.NAME
          const sourceID = sourse.STATUS_ID
          const workManagerDeals = this.dealsToSource.work.filter(e => sourse.STATUS_ID == e.sourse).length > 0 ? this.dealsToSource.work.filter(e => sourse.STATUS_ID == e.sourse).length : 0
          const closeManagerDeals = this.dealsToSource.close.filter(e => sourse.STATUS_ID == e.sourse).length > 0 ? this.dealsToSource.close.filter(e => sourse.STATUS_ID == e.sourse).length : 0
          const winManagerDeals = this.dealsToSource.win.filter(e => sourse.STATUS_ID == e.sourse).length > 0 ? this.dealsToSource.win.filter(e => sourse.STATUS_ID == e.sourse).length : 0
          const planToursManagerDeals = this.deals.toursPlan.filter(e => sourse.STATUS_ID == e.sourse).length > 0 ? this.deals.toursPlan.filter(e => sourse.STATUS_ID == e.sourse).length : 0
          const toursManagerDeals = this.deals.tours.filter(e => sourse.NAME == e.sourse).length > 0 ? this.deals.tours.filter(e => sourse.NAME == e.sourse).length : 0
          const doubleToursManagerDeals = this.deals.doubleTours.filter(e => sourse.NAME == e.sourse).length > 0 ? this.deals.doubleTours.filter(e => sourse.NAME == e.sourse).length : 0
          const dealsAllManagerDeals = Number(workManagerDeals) + Number(closeManagerDeals) + Number(winManagerDeals)
          if (workManagerDeals > 0 || closeManagerDeals > 0 || winManagerDeals > 0 || planToursManagerDeals > 0 || toursManagerDeals > 0 || doubleToursManagerDeals > 0 || dealsAllManagerDeals > 0) {
            const sourceItem = {
              manager: sourse.NAME,
              dealsAll: dealsAllManagerDeals,
              dealsWork: workManagerDeals,
              dealsClose: closeManagerDeals,
              toursPlan: planToursManagerDeals,
              toursWins: toursManagerDeals,
              DoubleTours: doubleToursManagerDeals,
              dealsWins: winManagerDeals,
            }
            if (sourceID == '2') {
              sourceItem['menu'] = false
              this.itemsSources.push(sourceItem)

              this.totalSourse.win += winManagerDeals
              this.totalSourse.work += workManagerDeals
              this.totalSourse.close += closeManagerDeals
              this.totalSourse.tours += toursManagerDeals
              this.totalSourse.doubleTours += doubleToursManagerDeals
              this.totalSourse.toursPlan += planToursManagerDeals
              this.totalSourse.totalDeals += dealsAllManagerDeals
            } else {
              if (sourceName.toLowerCase().includes('реклама') ||
                  sourceName.toLowerCase().includes('whatsapp') ||
                  sourceName.toLowerCase().includes('telegram') ||
                  sourceID == 'CALL' ||
                  sourceID == 'WEB' ||
                  sourceID == 'WEBFORM') {
                sourceItem['color'] = 'blue-grey lighten-3'
                promoSource.items.push(sourceItem)

                promoSource.dealsWins += winManagerDeals
                promoSource.dealsWork += workManagerDeals
                promoSource.dealsClose += closeManagerDeals
                promoSource.toursWins += toursManagerDeals
                promoSource.DoubleTours += doubleToursManagerDeals
                promoSource.toursPlan += planToursManagerDeals
                promoSource.dealsAll += dealsAllManagerDeals
              } else {
                sourceItem['color'] = 'blue-grey lighten-3'
                other.items.push(sourceItem)
                other.dealsWins += winManagerDeals
                other.dealsWork += workManagerDeals
                other.dealsClose += closeManagerDeals
                other.toursWins += toursManagerDeals
                other.DoubleTours += doubleToursManagerDeals
                other.toursPlan += planToursManagerDeals
                other.dealsAll += dealsAllManagerDeals
              }
            }
          }
        }
        this.itemsSources.push(promoSource)
        this.itemsSources.push(other)

        this.totalSourse.win += (promoSource.dealsWins + other.dealsWins)
        this.totalSourse.work += (promoSource.dealsWork + other.dealsWork)
        this.totalSourse.close += (promoSource.dealsClose + other.dealsClose)
        this.totalSourse.tours += (promoSource.toursWins + other.toursWins)
        this.totalSourse.doubleTours += (promoSource.DoubleTours + other.DoubleTours)
        this.totalSourse.toursPlan += (promoSource.toursPlan + other.toursPlan)
        this.totalSourse.totalDeals += (promoSource.dealsAll + other.dealsAll)
      } catch (e) {
        console.log(e)
      }
    },
    getReportPromo() {
      try {
        for (const promo of this.promoManagers) {
          const workManagerDeals = this.dealsToSource.work.filter(e => promo.ID == e.promo).length > 0 ? this.dealsToSource.work.filter(e => promo.ID == e.promo).length : 0
          const closeManagerDeals = this.dealsToSource.close.filter(e => promo.ID == e.promo).length > 0 ? this.dealsToSource.close.filter(e => promo.ID == e.promo).length : 0
          const winManagerDeals = this.dealsToSource.win.filter(e => promo.ID == e.promo).length > 0 ? this.dealsToSource.win.filter(e => promo.ID == e.promo).length : 0
          const planToursManagerDeals = this.deals.toursPlan.filter(e => promo.ID == e.promo).length > 0 ? this.deals.toursPlan.filter(e => promo.ID == e.promo).length : 0
          const toursManagerDeals = this.deals.tours.filter(e => promo.VALUE == e.promo).length > 0 ? this.deals.tours.filter(e => promo.VALUE == e.promo).length : 0
          const doubleToursManagerDeals = this.deals.doubleTours.filter(e => promo.VALUE == e.promo).length > 0 ? this.deals.doubleTours.filter(e => promo.VALUE == e.promo).length : 0
          const dealsAllManagerDeals = Number(workManagerDeals) + Number(closeManagerDeals) + Number(winManagerDeals)
          if (workManagerDeals > 0 || closeManagerDeals > 0 || winManagerDeals > 0 || planToursManagerDeals > 0 || toursManagerDeals > 0 || doubleToursManagerDeals > 0 || dealsAllManagerDeals > 0) {
            const promoItem = {
              manager: promo.VALUE,
              dealsAll: dealsAllManagerDeals,
              dealsWork: workManagerDeals,
              dealsClose: closeManagerDeals,
              toursPlan: planToursManagerDeals,
              toursWins: toursManagerDeals,
              DoubleTours: doubleToursManagerDeals,
              dealsWins: winManagerDeals,
            }
            this.itemsPromo.push(promoItem)

            this.totalPromo.win += winManagerDeals
            this.totalPromo.work += workManagerDeals
            this.totalPromo.close += closeManagerDeals
            this.totalPromo.tours += toursManagerDeals
            this.totalPromo.doubleTours += doubleToursManagerDeals
            this.totalPromo.toursPlan += planToursManagerDeals
            this.totalPromo.totalDeals += dealsAllManagerDeals
          }
        }
      } catch (e) {
        console.log(e)
      }
    },
    getButsToSource() {
      try {
        for (const buts of this.buts) {
          const workManagerDeals = this.dealsToSource.work.filter(e => buts.ID == e.buts).length > 0 ? this.dealsToSource.work.filter(e => buts.ID == e.buts).length : 0
          const closeManagerDeals = this.dealsToSource.close.filter(e => buts.ID == e.buts).length > 0 ? this.dealsToSource.close.filter(e => buts.ID == e.buts).length : 0
          const winManagerDeals = this.dealsToSource.win.filter(e => buts.ID == e.buts).length > 0 ? this.dealsToSource.win.filter(e => buts.ID == e.buts).length : 0
          const planToursManagerDeals = this.deals.toursPlan.filter(e => buts.ID == e.buts).length > 0 ? this.deals.toursPlan.filter(e => buts.ID == e.buts).length : 0
          const toursManagerDeals = this.deals.tours.filter(e => buts.VALUE == e.buts).length > 0 ? this.deals.tours.filter(e => buts.VALUE == e.buts).length : 0
          const doubleToursManagerDeals = this.deals.doubleTours.filter(e => buts.VALUE == e.buts).length > 0 ? this.deals.doubleTours.filter(e => buts.VALUE == e.buts).length : 0
          const dealsAllManagerDeals = Number(workManagerDeals) + Number(closeManagerDeals) + Number(winManagerDeals)
          if (workManagerDeals > 0 || closeManagerDeals > 0 || winManagerDeals > 0 || planToursManagerDeals > 0 || toursManagerDeals > 0 || doubleToursManagerDeals > 0 || dealsAllManagerDeals > 0) {
            const promoItem = {
              manager: buts.VALUE,
              color: 'blue-grey lighten-3',
              dealsAll: dealsAllManagerDeals,
              dealsWork: workManagerDeals,
              dealsClose: closeManagerDeals,
              toursPlan: planToursManagerDeals,
              toursWins: toursManagerDeals,
              DoubleTours: doubleToursManagerDeals,
              dealsWins: winManagerDeals,
            }
            const butsIndex = this.itemsSources.findIndex(e => e.manager == 'Бутсы')
            if (!this.itemsSources[butsIndex].items) this.itemsSources[butsIndex]['items'] = []
            this.itemsSources[butsIndex].items.push(promoItem)
            console.log(this.itemsSources, 'itemsSources')
          }
        }
      } catch (e) {
        console.log(e)
      }
    },
    async getDealsFunc(method, params, dealsWork = []) {
      try {
        const deals = []
        let isTrue = true
        while (isTrue) {
          //получаем все сделки, которые были в работе за отчетный период
          const allDealsData = await axios.post(this.url + method, params)
          await this.pause(250)
          for (const item of allDealsData.data.result.items) {
            if (!deals.includes(item['OWNER_ID'])) {
              deals.push(item['OWNER_ID'])
            }
            if (dealsWork.length > 0) {
              if (dealsWork.indexOf(item['OWNER_ID']) > -1) {
                dealsWork.splice(dealsWork.indexOf(item['OWNER_ID']), 1)
              }
            }
          }
          if (allDealsData.data.next) params.start = allDealsData.data.next
          else isTrue = false
        }
        return deals
      } catch (e) {
        console.log(e)
      }
    },
    async getDealsListFunc(method, params) {
      try {
        const deals = []
        const dealID = []
        let isTrue = true
        while (isTrue) {
          //получаем все сделки, которые были в работе за отчетный период
          const allDealsData = await axios.post(this.url + method, params)
          await this.pause(250)
          for (const item of allDealsData.data.result) {
            if (!dealID.includes(item['ID'])) {
              dealID.push(item['ID'])
              deals.push(item)
            }
          }
          if (allDealsData.data.next) params.start = allDealsData.data.next
          else isTrue = false
        }
        return deals
      } catch (e) {
        console.log(e)
      }
    },
    async bachFunc(items, type = 'type') {
      try {
        const batchParams = {
          'halt': 0,
          'cmd': {}
        }
        for (const index in items) {
          batchParams.cmd[index] = `crm.deal.get?id=${items[index]}`
          if (index % 50 == 0 || index == Number(items.length) - 1) {
            const batchResult = await axios.post(this.url + 'batch', batchParams)
            await this.pause(250)
            for (const dealIndex in batchResult.data.result.result) {
              this.deals[type].push({
                id: batchResult.data.result.result[dealIndex]['ID'],
                manager: batchResult.data.result.result[dealIndex]['ASSIGNED_BY_ID'],
                sourse: batchResult.data.result.result[dealIndex]['SOURCE_ID'],
                promo: batchResult.data.result.result[dealIndex]['UF_CRM_63944FE4EC0CD'],
                buts: batchResult.data.result.result[dealIndex]['UF_CRM_639FE7BFC4BBB']
              })
            }
            batchParams.cmd = {}
          }
        }
      } catch (e) {
        console.log(e)
      }
    },
    async getManager() {
      try {
        const usersData1 = await axios.post(this.url + 'user.get', {UF_DEPARTMENT: 5, ACTIVE: 'Y'})
        this.managers1 = usersData1.data.result
        await this.pause(50)
        const usersData2 = await axios.post(this.url + 'user.get', {UF_DEPARTMENT: 8, ACTIVE: 'Y'})
        this.managers2 = usersData2.data.result
      } catch (e) {
        console.log(e)
      }
    },
    async getSource() {
      const sourseData = await axios.post(this.url + 'crm.status.entity.items', {entityId: 'SOURCE'})
      this.sourse = sourseData.data.result
    },
    async getPromoManagers() {
      const promoManagersData = await axios.post(this.url + 'crm.deal.userfield.list', {filter: {ID: 464}})
      this.promoManagers = promoManagersData.data.result[0].LIST
    },
    async getButs() {
      const butsData = await axios.post(this.url + 'crm.deal.userfield.list', {filter: {ID: 532}})
      this.buts = butsData.data.result[0].LIST
      console.log(this.buts, 'buts')
    },
    async exportXls() {
      try {
        const workbook = new exceljs.Workbook();
        workbook.views = [
          {
            x: 0,
            y: 500,
            width: 29040,
            height: 15840,
            visibility: 'visible'
          }
        ]
        const worksheet = workbook.addWorksheet('Отчет по объектам');
        worksheet.columns = [
          {header: ' ', key: 'manager', width: 17, style: {numFmt: '#,0'}},
          {header: 'Общие кол Сделок', key: 'dealsAll', width: 17, style: {numFmt: '#,0'}},
          {header: 'В работе', key: 'dealsWork', width: 17, style: {numFmt: '#,0'}},
          {header: 'Закрыто', key: 'dealsClose', width: 17, style: {numFmt: '#,0'}},
          {header: 'Запланированы Туры', key: 'toursPlan', width: 17, style: {numFmt: '#,0'}},
          {header: 'Туры', key: 'toursWins', width: 17, style: {numFmt: '#,0'}},
          {header: 'Вторичные туры', key: 'DoubleTours', width: 17, style: {numFmt: '#,0'}},
          {header: 'Продажи', key: 'dealsWins', width: 17, style: {numFmt: '#,0'}},
        ];
        worksheet.addRow({
          manager: 'Менеджеры',
          dealsAll: ' ',
          dealsWork: ' ',
          toursPlan: ' ',
          toursWins: ' ',
          DoubleTours: ' ',
          dealsWins: ' '
        })
        for (const dealManager of this.itemsManager) {
          const managerRow = {
            manager: dealManager.manager,
            dealsAll: dealManager.dealsAll,
            dealsWork: dealManager.dealsWork,
            toursPlan: dealManager.toursPlan,
            toursWins: dealManager.toursWins,
            DoubleTours: dealManager.DoubleTours,
            dealsWins: dealManager.dealsWins
          }
          const styleManagerRow = worksheet.addRow(managerRow);
          /*  styleManagerRow.fill = {
              type: 'pattern',
              pattern: 'solid',
               fgColor: {argb: 'B9B7BD'},
            };*/
          styleManagerRow.border = {
            top: {style: 'thin'},
            left: {style: 'thin'},
            bottom: {style: 'thin'},
            right: {style: 'thin'}
          };
        }
        //total
        const managerRowTotal = {
          manager: 'Итого:  ',
          dealsAll: this.totalManager.totalDeals,
          dealsWork: this.totalManager.work,
          toursPlan: this.totalManager.toursPlan,
          toursWins: this.totalManager.tours,
          DoubleTours: this.totalManager.doubleTours,
          dealsWins: this.totalManager.win,
          dealsClose: this.totalManager.close
        }
        const styleManagerRowTotal = worksheet.addRow(managerRowTotal);
        /*styleManagerRowTotal.fill = {
          type: 'pattern',
          pattern: 'solid',
           fgColor: {argb: 'B9B7BD'},
        };*/
        styleManagerRowTotal.border = {
          top: {style: 'thin'},
          left: {style: 'thin'},
          bottom: {style: 'thin'},
          right: {style: 'thin'}
        };
        worksheet.addRow({
          manager: ' ',
          dealsAll: ' ',
          dealsWork: ' ',
          toursPlan: ' ',
          toursWins: ' ',
          DoubleTours: ' ',
          dealsWins: ' '
        })

        worksheet.addRow({
          manager: 'Промоутеры',
          dealsAll: ' ',
          dealsWork: ' ',
          toursPlan: ' ',
          toursWins: ' ',
          DoubleTours: ' ',
          dealsWins: ' '
        })
        for (const dealPromo of this.itemsPromo) {
          const promoRow = {
            manager: dealPromo.manager,
            dealsAll: dealPromo.dealsAll,
            dealsWork: dealPromo.dealsWork,
            toursPlan: dealPromo.toursPlan,
            toursWins: dealPromo.toursWins,
            DoubleTours: dealPromo.DoubleTours,
            dealsWins: dealPromo.dealsWins
          }
          const stylePromoRow = worksheet.addRow(promoRow);
          /*   stylePromoRow.fill = {
               type: 'pattern',
               pattern: 'solid',
               //fgColor: {argb: 'B9B7BD'},
             };*/
          stylePromoRow.border = {
            top: {style: 'thin'},
            left: {style: 'thin'},
            bottom: {style: 'thin'},
            right: {style: 'thin'}
          };
        }
        //total
        const dealPromoRowTotal = {
          manager: 'Итого:  ',
          dealsAll: this.totalPromo.totalDeals,
          dealsWork: this.totalPromo.work,
          toursPlan: this.totalPromo.toursPlan,
          toursWins: this.totalPromo.tours,
          DoubleTours: this.totalPromo.doubleTours,
          dealsWins: this.totalPromo.win,
          dealsClose: this.totalPromo.close
        }
        const stylePromoRowTotal = worksheet.addRow(dealPromoRowTotal);
        /*  stylePromoRowTotal.fill = {
            type: 'pattern',
            pattern: 'solid',
            // fgColor: {argb: 'B9B7BD'},
          };*/
        stylePromoRowTotal.border = {
          top: {style: 'thin'},
          left: {style: 'thin'},
          bottom: {style: 'thin'},
          right: {style: 'thin'}
        };
        worksheet.addRow({
          manager: ' ',
          dealsAll: ' ',
          dealsWork: ' ',
          toursPlan: ' ',
          toursWins: ' ',
          DoubleTours: ' ',
          dealsWins: ' '
        })

        worksheet.addRow({
          manager: 'Источники',
          dealsAll: ' ',
          dealsWork: ' ',
          toursPlan: ' ',
          toursWins: ' ',
          DoubleTours: ' ',
          dealsWins: ' '
        })
        for (const dealSource of this.itemsSources) {
          const sourceRow = {
            manager: dealSource.manager,
            dealsAll: dealSource.dealsAll,
            dealsWork: dealSource.dealsWork,
            toursPlan: dealSource.toursPlan,
            toursWins: dealSource.toursWins,
            DoubleTours: dealSource.DoubleTours,
            dealsWins: dealSource.dealsWins
          }
          const stylesourceRow = worksheet.addRow(sourceRow);
          /*  stylesourceRow.fill = {
              type: 'pattern',
              pattern: 'solid',
              // fgColor: {argb: 'B9B7BD'},
            };*/
          stylesourceRow.border = {
            top: {style: 'thin'},
            left: {style: 'thin'},
            bottom: {style: 'thin'},
            right: {style: 'thin'}
          };
          if (dealSource.items.length > 0) {
            for (const buts of dealSource.items) {
              const butsRow = {
                manager: buts.manager,
                dealsAll: buts.dealsAll,
                dealsWork: buts.dealsWork,
                toursPlan: buts.toursPlan,
                toursWins: buts.toursWins,
                DoubleTours: buts.DoubleTours,
                dealsWins: buts.dealsWins
              }
              const styleButsRow = worksheet.addRow(butsRow);
              styleButsRow.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: {argb: 'B9B7BD'},
              };
              styleButsRow.border = {
                top: {style: 'thin'},
                left: {style: 'thin'},
                bottom: {style: 'thin'},
                right: {style: 'thin'}
              };
            }
          }
        }
        //total
        const dealSourceRowTotal = {
          manager: 'Итого:  ',
          dealsAll: this.totalSourse.totalDeals,
          dealsWork: this.totalSourse.work,
          toursPlan: this.totalSourse.toursPlan,
          toursWins: this.totalSourse.tours,
          DoubleTours: this.totalSourse.doubleTours,
          dealsWins: this.totalSourse.win,
          dealsClose: this.totalSourse.close
        }
        const styleSourceRowTotal = worksheet.addRow(dealSourceRowTotal);
        /*styleSourceRowTotal.fill = {
          type: 'pattern',
          pattern: 'solid',
          // fgColor: {argb: 'B9B7BD'},
        };*/
        styleSourceRowTotal.border = {
          top: {style: 'thin'},
          left: {style: 'thin'},
          bottom: {style: 'thin'},
          right: {style: 'thin'}
        };
        worksheet.addRow({
          manager: ' ',
          dealsAll: ' ',
          dealsWork: ' ',
          toursPlan: ' ',
          toursWins: ' ',
          DoubleTours: ' ',
          dealsWins: ' '
        })

        const buffer = await workbook.xlsx.writeBuffer();
        const fileType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
        const blob = new Blob([buffer], {type: fileType});
        fileWrite(blob, 'report.xlsx');
      } catch (e) {
        console.log(e)
      }
    },
    onResize() {
      this.isMobile = window.innerWidth < 600
    },
  },
  beforeDestroy() {
    if (typeof window === 'undefined') return
    window.removeEventListener('resize', this.onResize, {passive: true})
  },
  async mounted() {
    this.getManager();
    await this.pause(100)
    this.getSource();
    await this.pause(100)
    this.getPromoManagers();
    await this.pause(100)
    this.getButs();
    const date = new Date()
    const month = date.getMonth() + 1
    const year = date.getFullYear()
    this.dates.push(`${year}-${month}-01`)
    this.dates.push(`${year}-${month}-${this.monthMatrix[month]}`)
    this.onResize()
    window.addEventListener('resize', this.onResize, {passive: true})
    this.loading = false
  }
};
</script>

<style scoped>
.divider {
  margin-bottom: 50px;
}
</style>
