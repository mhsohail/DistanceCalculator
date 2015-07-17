using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.Linq;
using System.Net;
using System.Web;
using System.Web.Mvc;
using DistanceCalculator.Models;
using DistanceCalculator.ViewModels;

namespace DistanceCalculator.Controllers
{
    public class HomeIndexViewModelsController : Controller
    {
        private DistanceCalculatorContext db = new DistanceCalculatorContext();

        // GET: HomeIndexViewModels
        public ActionResult Index()
        {
            return View(db.HomeIndexViewModels.ToList());
        }

        // GET: HomeIndexViewModels/Details/5
        public ActionResult Details(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            HomeIndexViewModel homeIndexViewModel = db.HomeIndexViewModels.Find(id);
            if (homeIndexViewModel == null)
            {
                return HttpNotFound();
            }
            return View(homeIndexViewModel);
        }

        // GET: HomeIndexViewModels/Create
        public ActionResult Create()
        {
            return View();
        }

        // POST: HomeIndexViewModels/Create
        // To protect from overposting attacks, please enable the specific properties you want to bind to, for 
        // more details see http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create(string id)
        {
            if (ModelState.IsValid)
            {
                //db.HomeIndexViewModels.Add(homeIndexViewModel);
                db.SaveChanges();
                return RedirectToAction("Index");
            }

            return View();
        }

        // GET: HomeIndexViewModels/Edit/5
        public ActionResult Edit(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            HomeIndexViewModel homeIndexViewModel = db.HomeIndexViewModels.Find(id);
            if (homeIndexViewModel == null)
            {
                return HttpNotFound();
            }
            return View(homeIndexViewModel);
        }

        // POST: HomeIndexViewModels/Edit/5
        // To protect from overposting attacks, please enable the specific properties you want to bind to, for 
        // more details see http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit([Bind(Include = "Id,ExcelFile")] HomeIndexViewModel homeIndexViewModel)
        {
            if (ModelState.IsValid)
            {
                db.Entry(homeIndexViewModel).State = EntityState.Modified;
                db.SaveChanges();
                return RedirectToAction("Index");
            }
            return View(homeIndexViewModel);
        }

        // GET: HomeIndexViewModels/Delete/5
        public ActionResult Delete(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            HomeIndexViewModel homeIndexViewModel = db.HomeIndexViewModels.Find(id);
            if (homeIndexViewModel == null)
            {
                return HttpNotFound();
            }
            return View(homeIndexViewModel);
        }

        // POST: HomeIndexViewModels/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public ActionResult DeleteConfirmed(int id)
        {
            HomeIndexViewModel homeIndexViewModel = db.HomeIndexViewModels.Find(id);
            db.HomeIndexViewModels.Remove(homeIndexViewModel);
            db.SaveChanges();
            return RedirectToAction("Index");
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                db.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}
