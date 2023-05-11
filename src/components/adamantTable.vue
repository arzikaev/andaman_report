<template>
  <v-simple-table>
    <template v-slot:default>
      <thead>
      <tr class="text-left">
        <th :class="elem.class"
            v-for="elem in headers"
            :key="elem.value"
        >
          {{ elem.text }}
        </th>
      </tr>
      </thead>
      <tbody>
      <template v-for="item in items">
        <tr class="text-left"
            :key="item.manager"
        >
          <td class="text-left td">{{ item.manager }}
            <v-btn
                icon
                @click="item.menu = true"
                v-if="item.menu === false"
            >+
            </v-btn>
            <v-btn
                @click="item.menu= false"
                v-if="item.menu === true"
                icon
            >-
            </v-btn>
          </td>
          <td class="text-left td dealsAll">{{ item.dealsAll }}</td>
          <td class="text-left td">{{ item.dealsWork }}</td>
          <td class="text-left td">{{ item.dealsClose }}</td>
          <td class="text-left td">{{ item.toursPlan }}</td>
          <td class="text-left td toursWins">{{ item.toursWins }}</td>
          <td class="text-left td">{{ item.DoubleTours }}</td>
          <td class="text-left td dealsWins">{{ item.dealsWins }}</td>
        </tr>
        <template v-if="item.menu === true">
          <template v-for="itemButs in item.items">
            <tr class="text-left"
                :class="itemButs.color"
                :key="itemButs.manager"
            >
              <td class="text-left td">{{ itemButs.manager }}</td>
              <td class="text-left td dealsAll">{{ itemButs.dealsAll }}</td>
              <td class="text-left td">{{ itemButs.dealsWork }}</td>
              <td class="text-left td">{{ itemButs.dealsClose }}</td>
              <td class="text-left td">{{ itemButs.toursPlan }}</td>
              <td class="text-left td toursWins">{{ itemButs.toursWins }}</td>
              <td class="text-left td">{{ itemButs.DoubleTours }}</td>
              <td class="text-left td dealsWins">{{ itemButs.dealsWins }}</td>
            </tr>
          </template>
        </template>
      </template>
      <tr :class="deals.color">
        <td class="text-left td">Итого:</td>
        <td class="text-left td dealsAll">{{ deals.totalDeals }}</td>
        <td class="text-left td">{{ deals.work }}</td>
        <td class="text-left td">{{ deals.close }}</td>
        <td class="text-left td">{{ deals.toursPlan }}</td>
        <td class="text-left td toursWins">{{ deals.tours }}</td>
        <td class="text-left td">{{ deals.doubleTours }}</td>
        <td class="text-left td dealsWins">{{ deals.win }}</td>
      </tr>
      </tbody>
    </template>
  </v-simple-table>
</template>

<script>

export default {
  name: "adamantTable",
  props: ['items', 'menu', 'deals'],
  data() {
    return {
      headers: [
        {
          text: ' ',
          align: 'start',
          sortable: false,
          value: 'manager',
          class: 'text-left th'
        },
        {
          text: 'Общие кол Сделок',
          sortable: true,
          value: 'dealsAll',
          class: 'text-left th dealsAll'
        },
        {
          text: 'В работе',
          sortable: true,
          value: 'dealsWork',
          class: 'text-left th'
        },
        {
          text: 'Закрыто',
          sortable: true,
          value: 'dealsClose',
          class: 'text-left th'
        },
        {
          text: 'Запланированы Туры',
          sortable: true,
          value: 'toursPlan',
          class: 'text-left th'
        },
        {
          text: 'Туры',
          sortable: true,
          value: 'toursWins',
          class: 'text-left th toursWins'
        },
        {
          text: 'Вторичные туры',
          sortable: true,
          value: 'DoubleTours',
          class: 'text-left th'
        },
        {
          text: 'Продажи',
          sortable: true,
          value: 'dealsWins',
          class: 'text-left th dealsWins'
        }
      ]
    }
  }
}
</script>

<style scoped>

.td {
  border: 1px solid black;
}

.th {
  border: 1px solid black;
}

.dealsWins {
  background-color: yellowgreen;
}

.toursWins {
  background-color: lightgreen;
}

.dealsAll {
  background-color: lightsalmon;
}
</style>