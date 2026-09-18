#!/usr/bin/env python3
"""Lokální demonstrátor systému řízení odchylek. Spusťte: python3 server.py"""
import json, sqlite3, os
from datetime import datetime
from http.server import ThreadingHTTPServer, SimpleHTTPRequestHandler

ROOT = os.path.dirname(os.path.abspath(__file__))
DB = os.path.join(ROOT, "odchylky-demo.sqlite3")
STEPS = ["Nová", "Posouzení rizika", "Vyšetřování", "CAPA", "Ověření účinnosti", "Schválení uzavření", "Uzavřena"]

def now(): return datetime.now().strftime("%Y-%m-%d %H:%M")
def con():
    c = sqlite3.connect(DB); c.row_factory = sqlite3.Row; return c
def setup():
    c = con()
    c.executescript("""
      CREATE TABLE IF NOT EXISTS packages(id INTEGER PRIMARY KEY, ref TEXT, source TEXT, location TEXT, auditor TEXT, deadline TEXT, created_at TEXT);
      CREATE TABLE IF NOT EXISTS deviations(id INTEGER PRIMARY KEY, package_id INTEGER, title TEXT, scope TEXT, source TEXT, location TEXT, product_regime TEXT, process TEXT, risk INTEGER, status TEXT, owner TEXT, deadline TEXT, stock_state TEXT, fmea_required INTEGER, fmea_decision TEXT, description TEXT, created_at TEXT);
      CREATE TABLE IF NOT EXISTS audit(id INTEGER PRIMARY KEY, deviation_id INTEGER, action TEXT, detail TEXT, actor TEXT, created_at TEXT);
      CREATE TABLE IF NOT EXISTS capa(id INTEGER PRIMARY KEY, deviation_id INTEGER, type TEXT, description TEXT, owner TEXT, deadline TEXT, status TEXT, evidence TEXT, effectiveness TEXT, created_at TEXT);
      CREATE TABLE IF NOT EXISTS fmea(id INTEGER PRIMARY KEY, process TEXT, failure_mode TEXT, effect_text TEXT, cause_text TEXT, prevention TEXT, detection TEXT, severity INTEGER, occurrence INTEGER, detection_score INTEGER, owner TEXT, status TEXT, updated_at TEXT);
    """)
    if c.execute("SELECT count(*) FROM deviations").fetchone()[0] == 0:
      c.execute("INSERT INTO packages(ref,source,location,auditor,deadline,created_at) VALUES(?,?,?,?,?,?)", ("IK-2026-041","Interní kontrola","Centrální sklad","Jana Nováková","2026-10-18",now()))
      pid=c.execute("SELECT last_insert_rowid()").fetchone()[0]
      seed=[("Chybí záznam o kontrole teploty","Proces","Teplotní režim",4,"Posouzení rizika","Vedoucí skladu","2026-09-25","—",1),("Neúplné značení vyčleněné zásoby","Produkt i proces","Krmiva",3,"Vyšetřování","QA","2026-09-28","Blokováno / karanténa",1),("Pracovní postup není aktuální","Proces","Řízená dokumentace",2,"CAPA","QA","2026-10-10","—",0)]
      for title,scope,process,risk,status,owner,deadline,stock,fmea in seed:
        c.execute("INSERT INTO deviations(package_id,title,scope,source,location,product_regime,process,risk,status,owner,deadline,stock_state,fmea_required,description,created_at) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", (pid,title,scope,"Interní kontrola","Centrální sklad","Krmiva" if "Produkt" in scope else "",process,risk,status,owner,deadline,stock,fmea,"Ukázkový nález vytvořený pro ověření chování systému.",now()))
        did=c.execute("SELECT last_insert_rowid()").fetchone()[0]; c.execute("INSERT INTO audit(deviation_id,action,detail,actor,created_at) VALUES(?,?,?,?,?)",(did,"Založeno","Nález založen z balíčku IK-2026-041","Systém",now()))
    if c.execute("SELECT count(*) FROM fmea").fetchone()[0] == 0:
      c.execute("INSERT INTO fmea(process,failure_mode,effect_text,cause_text,prevention,detection,severity,occurrence,detection_score,owner,status,updated_at) VALUES(?,?,?,?,?,?,?,?,?,?,?,?)",("Teplotní režim","Chybějící záznam teploty","Nelze doložit podmínky skladování","Neúplná kontrola směny","Školení a kontrolní seznam","Denní kontrola QA",5,2,3,"QA","Aktivní",now()))
    c.commit(); c.close()
def rows(sql,args=()):
    c=con(); result=[dict(x) for x in c.execute(sql,args).fetchall()]; c.close(); return result
def write(sql,args=()):
    c=con(); c.execute(sql,args); c.commit(); ident=c.execute("SELECT last_insert_rowid()").fetchone()[0]; c.close(); return ident
def audit(did,action,detail,actor="Demo uživatel"): write("INSERT INTO audit(deviation_id,action,detail,actor,created_at) VALUES(?,?,?,?,?)",(did,action,detail,actor,now()))
def deviation(did):
    data=rows("SELECT d.*, p.ref package_ref FROM deviations d LEFT JOIN packages p ON p.id=d.package_id WHERE d.id=?",(did,)); return data[0] if data else None
class App(SimpleHTTPRequestHandler):
  def send_json(self,obj,status=200):
    payload=json.dumps(obj,ensure_ascii=False).encode(); self.send_response(status); self.send_header("Content-Type","application/json; charset=utf-8"); self.send_header("Content-Length",str(len(payload))); self.end_headers(); self.wfile.write(payload)
  def body(self): return json.loads(self.rfile.read(int(self.headers.get("Content-Length",0))) or b"{}")
  def do_GET(self):
    if self.path.startswith("/api/overview"):
      ds=rows("SELECT * FROM deviations ORDER BY CASE status WHEN 'Uzavřena' THEN 9 ELSE 0 END, deadline, id DESC")
      self.send_json({"deviations":ds,"packages":rows("SELECT p.*,count(d.id) findings FROM packages p LEFT JOIN deviations d ON d.package_id=p.id GROUP BY p.id ORDER BY p.id DESC"),"steps":STEPS}); return
    if self.path.startswith("/api/deviations/"):
      did=int(self.path.split("/")[3]); self.send_json({"deviation":deviation(did),"audit":rows("SELECT * FROM audit WHERE deviation_id=? ORDER BY id DESC",(did,)),"capa":rows("SELECT * FROM capa WHERE deviation_id=? ORDER BY id DESC",(did,)),"steps":STEPS}); return
    if self.path.startswith("/api/packages/"):
      pid=int(self.path.split("/")[3]); package=rows("SELECT * FROM packages WHERE id=?",(pid,))
      if not package: return self.send_json({"error":"Nenalezeno"},404)
      self.send_json({"package":package[0],"deviations":rows("SELECT * FROM deviations WHERE package_id=? ORDER BY id",(pid,))}); return
    if self.path.startswith("/api/fmea"):
      self.send_json({"records":rows("SELECT *, severity*occurrence*detection_score rpn FROM fmea ORDER BY rpn DESC")}); return
    if self.path.startswith("/api/tasks"):
      self.send_json({"tasks":rows("SELECT d.id,'Revize FMEA' title,d.owner,d.deadline,d.status FROM deviations d WHERE d.fmea_required=1 AND (d.fmea_decision IS NULL OR d.fmea_decision='') UNION ALL SELECT c.id,'CAPA: '||c.description,c.owner,c.deadline,c.status FROM capa c")}); return
    return super().do_GET()
  def do_POST(self):
    try: data=self.body()
    except Exception: return self.send_json({"error":"Neplatná data"},400)
    if self.path=="/api/deviations":
      risk=int(data.get("risk",3)); fmea=1 if risk>=4 or data.get("scope") in ("Produkt","Produkt i proces") else 0
      did=write("INSERT INTO deviations(package_id,title,scope,source,location,product_regime,process,risk,status,owner,deadline,stock_state,fmea_required,description,created_at) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",(None,data["title"],data["scope"],data["source"],data.get("location",""),data.get("product_regime",""),data.get("process",""),risk,"Nová",data.get("owner","QA"),data.get("deadline",""),data.get("stock_state",""),fmea,data.get("description",""),now()))
      audit(did,"Založeno","Nový podnět založen."); self.send_json({"id":did,"message":"Podnět byl založen."},201); return
    if self.path=="/api/packages":
      pid=write("INSERT INTO packages(ref,source,location,auditor,deadline,created_at) VALUES(?,?,?,?,?,?)",(data["ref"],data["source"],data.get("location",""),data.get("auditor",""),data.get("deadline",""),now()))
      ids=[]
      for item in data["findings"]:
        risk=int(item.get("risk",3)); fmea=1 if risk>=4 or item.get("scope") in ("Produkt","Produkt i proces") else 0
        did=write("INSERT INTO deviations(package_id,title,scope,source,location,product_regime,process,risk,status,owner,deadline,stock_state,fmea_required,description,created_at) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",(pid,item["title"],item["scope"],data["source"],data.get("location",""),item.get("product_regime",""),item.get("process",""),risk,"Nová",item.get("owner","QA"),item.get("deadline",data.get("deadline","")),"",fmea,item.get("description",""),now())); audit(did,"Založeno","Založeno z balíčku "+data["ref"],"Systém"); ids.append(did)
      self.send_json({"id":pid,"deviations":ids,"message":"Balíček a "+str(len(ids))+" odchylky byly založeny."},201); return
    if self.path=="/api/fmea":
      fid=write("INSERT INTO fmea(process,failure_mode,effect_text,cause_text,prevention,detection,severity,occurrence,detection_score,owner,status,updated_at) VALUES(?,?,?,?,?,?,?,?,?,?,?,?)",(data["process"],data["failure_mode"],data.get("effect_text",""),data.get("cause_text",""),data.get("prevention",""),data.get("detection",""),int(data["severity"]),int(data["occurrence"]),int(data["detection_score"]),data.get("owner","QA"),"Aktivní",now()))
      self.send_json({"id":fid},201); return
    if self.path.startswith("/api/deviations/") and self.path.endswith("/action"):
      did=int(self.path.split("/")[3]); d=deviation(did)
      if not d: return self.send_json({"error":"Nenalezeno"},404)
      action=data.get("action")
      if action=="fmea":
        decision=data.get("decision")
        if decision not in ("FMEA aktualizována","FMEA beze změny"): return self.send_json({"error":"Vyberte rozhodnutí FMEA."},400)
        write("UPDATE deviations SET fmea_decision=? WHERE id=?",(decision,did)); audit(did,"Revize FMEA",decision); return self.send_json({"message":"Rozhodnutí FMEA uloženo."})
      try: next_status=STEPS[STEPS.index(d["status"])+1]
      except (ValueError,IndexError): return self.send_json({"error":"Případ již nelze posunout."},400)
      if next_status=="Uzavřena" and d["fmea_required"] and not d["fmea_decision"]: return self.send_json({"error":"Nelze uzavřít: čeká se na povinné rozhodnutí o revizi FMEA."},409)
      write("UPDATE deviations SET status=? WHERE id=?",(next_status,did)); audit(did,"Workflow",d["status"]+" → "+next_status); return self.send_json({"message":"Případ posunut do stavu: "+next_status,"status":next_status})
    if self.path.startswith("/api/deviations/") and self.path.endswith("/capa"):
      did=int(self.path.split("/")[3]); cid=write("INSERT INTO capa(deviation_id,type,description,owner,deadline,status,evidence,effectiveness,created_at) VALUES(?,?,?,?,?,?,?,?,?)",(did,data.get("type","Nápravné"),data["description"],data["owner"],data.get("deadline",""),"Otevřené","","",now()))
      audit(did,"CAPA založeno","Opatření #"+str(cid)+": "+data["description"]); self.send_json({"id":cid},201); return
    self.send_json({"error":"Nenalezeno"},404)
if __name__=="__main__":
  setup(); print("Demonstrátor běží na http://localhost:8765"); ThreadingHTTPServer(("127.0.0.1",8765),App).serve_forever()
