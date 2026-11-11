(function () {
  var slides = Array.prototype.slice.call(document.querySelectorAll(".slide"));
  var links = Array.prototype.slice.call(document.querySelectorAll(".toc a"));
  if (!slides.length) return;

  function current() {
    var line = window.scrollY + 140;
    var index = 0;
    for (var i = 0; i < slides.length; i++) {
      if (slides[i].offsetTop <= line) index = i;
    }
    return index;
  }

  function mark() {
    var index = current();
    links.forEach(function (link, i) {
      if (i === index) link.setAttribute("aria-current", "true");
      else link.removeAttribute("aria-current");
    });
  }

  function go(step) {
    var index = Math.max(0, Math.min(slides.length - 1, current() + step));
    slides[index].scrollIntoView({ behavior: "smooth", block: "start" });
  }

  document.addEventListener("keydown", function (event) {
    var tag = event.target && event.target.tagName;
    if (tag === "INPUT" || tag === "TEXTAREA") return;
    if (event.key === "j") { event.preventDefault(); go(1); }
    if (event.key === "k") { event.preventDefault(); go(-1); }
  });
  document.addEventListener("scroll", mark, { passive: true });
  mark();
})();
