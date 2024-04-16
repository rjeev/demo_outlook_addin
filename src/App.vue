<template>
  <div id="app">
    <div class="content">
      <div class="content-header">
        <div class="padding">
          <h1>Welcome</h1>
        </div>
      </div>
      <div class="content-main">

        <div class="content" v-if="subject">
          <div><b>Subject:</b> {{ subject }}</div>
          <div><b>Sender:</b> {{ sender }}</div>

        </div>
        <button class="myBtn" @click="handleClick">Run</button>
      </div>
    </div>
  </div>

</template>

<script>
export default {
  name: 'App',
  data() {
    return {
      subject: '',
      sender: '',

    }
  },
  methods: {
    handleClick() {
      this.subject = window.Office.context.mailbox.item.subject;
      this.sender = window.Office.context.mailbox.item.from;

    }
  }
};
</script>

<style>
.content-header {
  background: #2a8dd4;
  color: #fff;
  position: absolute;
  top: 0;
  left: 0;
  width: 100%;
  height: 80px;
  overflow: hidden;
}

.content-main {
  background: #fff;
  position: fixed;
  top: 80px;
  left: 0;
  right: 0;
  bottom: 0;
  overflow: auto;
}

.content {
  margin-top: 40px;
  margin-bottom: 40px;
  padding: 5px
}

.myBtn {
  padding: 5px;
  background-color: #2a8dd4;

  margin-left: 50%;
  margin-top: 50%;
}

.padding {
  padding: 15px;
}
</style>