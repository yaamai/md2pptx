# md2pptx

## Todo
- [x] pyodide + python-pptx PoC
  ```
  ./bin/pyodide mkpkg python-pptx
  ./bin/pyodide mkpkg marko
  # use github source due to setup.py not included in pypi source tarball.
  cat <<EOF > packages/marko/meta.yaml
  package:
    name: marko
    version: 1.0.1
  source:
    sha256: 76f79281e4503a07fea72cc7119d45690e7c2e368362b4568a764ccddcdefbd9
    url: https://github.com/frostming/marko/archive/v1.0.1.zip
  test:
    imports:
    - marko
  EOF
  PYODIDE_PACKAGES="micropip,lxml,python-pptx,marko" make
  ```
- [x] markdown parse
  - [x] marko
    ```
    import io
    from marko import Markdown
    from marko import Parser
    from marko.html_renderer import HTMLRenderer
    from marko.renderer import Renderer
    from pptx import Presentation

    """
    - heading(level=1) starts new slide
    - horizonal line starts new slide
    - heading(level>1) as
    """

    class PptxRenderer(Renderer):

        def render_document(self, element):
            self.pres = Presentation()
            self.render_children(element)

            buf = io.BytesIO()
            self.pres.save(buf)
            return buf

        def render_heading(self, element):
            print(element)

        def render_children(self, element):
            print(element)
            [self.render(child) for child in element.children]


    # md = Markdown(parser=Parser, renderer=HTMLRenderer)
    md = Markdown(parser=Parser, renderer=PptxRenderer)
    print(md("# aaa"))
    ```
- [ ] markdown to pptx logics
  - [ ] borrow from pandoc logic
- [ ] styling from marp
  - [ ] style template (slide-master)
  - [ ] precious styling with metadata (commonmark experimental attrs.)
    - https://pandoc.org/MANUAL.html#extension-header_attributes
- [ ] live preview (pptx->img?)
  - [ ] libreoffice?
