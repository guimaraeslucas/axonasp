<script runat="server" language="JScript">
		class Animal {
			constructor(name) {
				this.name = name;
			}
			speak() {
				return this.name + " makes a noise";
			}
		}
		class Dog extends Animal {
			constructor(name) {
				super(name);
			}
			speak() {
				return this.name + " barks";
			}
		}
		var d = new Dog("Mitzie");
		Response.Write(d.speak());
</script>
