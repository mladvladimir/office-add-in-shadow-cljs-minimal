(ns taskpane
  (:require
    ["office" :as office]))

(set! *warn-on-infer* true)

(enable-console-print!)

(defn hello [] "hello There from taskpane!")
(println (hello))

(defn run []
  (try
    (.run js/Excel
          (fn [^js context]
            ;Insert your Excel code here
            (let [range (.getSelectedRange (.-workbook context))]
              ;Read the range addres
              (.load range "address")
              ;Update the fill color
              (set! (.-color (.-fill (.-format range))) "yellow")
              (.then (.sync context)))))
    (catch js/Error err
      (js/console.log (ex-cause err)))))

(.onReady js/Office
          (fn [info]
            (when (= (.-host info) (.-Excel (.-HostType js/Office)))
              (do
                (set! (.-display (.-style (.getElementById js/document "sideload-msg"))) "none")
                (set! (.-display (.-style (.getElementById js/document "app-body"))) "flex")
                (.addEventListener (.getElementById js/document "run") "click" run)))))




