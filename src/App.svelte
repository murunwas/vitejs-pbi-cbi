<script>
import {onMount} from "svelte"
import {initPBI} from "./lib/bi"

let powerBI;

onMount(()=>{
  powerBI=initPBI()
})
</script>

<main>
<button on:click={()=>powerBI.print()}>Print</button>
<section id="container" style="height: 750px"/>
</main>