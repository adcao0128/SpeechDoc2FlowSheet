<template>
  <div id="app">
    <DocDownload />
  </div>
</template>

<script>
import DocDownload from './components/DocDownload.vue'

export default {
  name: 'App',
  components: {
    DocDownload
  }
}
</script>

<style>
#app {
  font-family: Avenir, Helvetica, Arial, sans-serif;
  -webkit-font-smoothing: antialiased;
  -moz-osx-font-smoothing: grayscale;
  text-align: center;
  color: #2c3e50;
  margin-top: 60px;
}
</style>
