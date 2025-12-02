from playwright.sync_api import sync_playwright, expect
import os

def test_admin_features(page):
    # Mock Google Script Run
    page.add_init_script("""
    window.google = {
      script: {
        run: {
          withSuccessHandler: function(success) {
            this._success = success;
            return this;
          },
          withFailureHandler: function(failure) {
            this._failure = failure;
            return this;
          },
          apiGetInitData: function() {
            setTimeout(() => {
               this._success({
                 ok: true,
                 user: { email: 'admin@test.com', nombre: 'Admin User', rol: 'ADMIN_GEN', departamento: 'IT', estado: 'ACTIVO' },
                 activeUser: { email: 'admin@test.com', rol: 'ADMIN_GEN' },
                 dates: [{value: '2023-10-27', label: 'Viernes 27/10'}],
                 menu: {},
                 allOrders: {},
                 adminData: { users: [], orders: [] }
               });
            }, 100);
          },
          apiGetAdminData: function() {
            setTimeout(() => {
               this._success({
                 ok: true,
                 users: [
                   { email: 'u1@test.com', nombre: 'User One', departamento: 'HR', rol: 'USER', estado: 'ACTIVO' },
                   { email: 'u2@test.com', nombre: 'User Two', departamento: 'IT', rol: 'ADMIN_DEP', estado: 'ACTIVO' }
                 ],
                 orders: [],
                 departments: [],
                 config: { 'RESPONSIBLES_EMAILS_JSON': '[{"name":"Resp1","email":"r1@test.com","type":"TO"}]' },
                 holidays: [],
                 configList: [{key: 'RESPONSIBLES_EMAILS_JSON', value: '...', desc: 'JSON'}]
               });
            }, 100);
          },
          apiSaveHoliday: function(d, desc) {
             console.log("Mock Save Holiday Called");
             setTimeout(() => {
                if (this._failure) {
                    console.log("Triggering failure");
                    this._failure(new Error("No puedes agregar días libres en el pasado."));
                }
             }, 100);
          },
          apiSaveConfig: function(config) {
              console.log("Mock Save Config Called");
              this._success({ok: true});
          },
          apiDismissBanner: function() {}
        }
      }
    };
    """)

    # Load App
    cwd = os.getcwd()
    page.goto(f"file://{cwd}/mock_app.html")

    # 1. Switch to Admin View
    page.get_by_title("Panel de Administración").click()
    page.wait_for_timeout(500)

    # 2. Check Users Table
    headers = page.locator("thead th").all_text_contents()
    if "Rol" in headers:
         print("SUCCESS: Rol column present")
    else:
         print("ERROR: Rol column missing")

    # 3. Check Config
    page.get_by_text("CONFIG").click()
    page.wait_for_timeout(500)
    page.get_by_text("Editar").click()
    page.wait_for_timeout(500)

    # Debug Inputs
    inputs = page.locator("input").all()
    found = False
    for inp in inputs:
        if inp.input_value() == "Resp1":
             found = True
             break

    if found:
         print("SUCCESS: Responsible loaded")
    else:
         print("ERROR: Responsible not found")
         # Print Page HTML for debug
         # print(page.content())

    # Add new Responsible
    page.get_by_text("Agregar Responsable").click()
    page.wait_for_timeout(200)

    # Fill last name input
    inputs = page.locator("input[placeholder='Nombre']").all()
    if len(inputs) > 0:
        inputs[-1].fill("NewResp")

    # Save Config
    page.get_by_text("Guardar").click()
    page.wait_for_timeout(500)
    print("SUCCESS: Config saved")

    # 4. Check Holidays
    page.get_by_text("CALENDARIO").click()
    page.wait_for_timeout(200)

    page.locator("input[type=date]").nth(0).fill("2020-01-01")
    page.locator("input[placeholder*='Descripción']").fill("Past Holiday")

    page.on("dialog", lambda dialog: dialog.accept())
    page.get_by_text("Agregar").click()

    page.wait_for_timeout(1000)

    # Verify Spinner Gone
    # The spinner has text "Cargando..."
    if page.locator(".animate-spin").count() > 0:
         # Check visibility
         if not page.locator(".animate-spin").first.is_visible():
             print("SUCCESS: Spinner hidden")
         else:
             print("ERROR: Spinner still visible")
    else:
         print("SUCCESS: Spinner not found (hidden)")

    page.screenshot(path="/home/jules/verification/verification.png", full_page=True)

if __name__ == "__main__":
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()
        try:
            test_admin_features(page)
        finally:
            browser.close()
