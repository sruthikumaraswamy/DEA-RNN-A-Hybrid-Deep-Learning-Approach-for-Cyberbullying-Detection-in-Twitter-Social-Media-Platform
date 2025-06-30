
from django.db.models import  Count, Avg
from django.shortcuts import render, redirect
from django.db.models import Count
from django.db.models import Q
import datetime
import xlwt
from django.http import HttpResponse


import re
import string
import pandas as pd
from wordcloud import WordCloud, STOPWORDS
from sklearn.feature_extraction.text import CountVectorizer
from sklearn.metrics import accuracy_score, confusion_matrix, classification_report
from sklearn.metrics import accuracy_score
from sklearn.metrics import f1_score
from sklearn.tree import DecisionTreeClassifier
from sklearn.ensemble import VotingClassifier
# Create your views here.
from Remote_User.models import ClientRegister_Model,Cyberbullying_Detection_Type,detection_ratio,detection_accuracy


def serviceproviderlogin(request):
    if request.method  == "POST":
        admin = request.POST.get('username')
        password = request.POST.get('password')
        if admin == "Admin" and password =="Admin":
            return redirect('View_Remote_Users')

    return render(request,'SProvider/serviceproviderlogin.html')

def Find_Predicted_Cyberbullying_Detection_Ratio(request):
    detection_ratio.objects.all().delete()
    ratio = ""
    kword = 'not_cyberbullying'
    print(kword)
    obj = Cyberbullying_Detection_Type.objects.all().filter(Q(Prediction=kword))
    obj1 = Cyberbullying_Detection_Type.objects.all()
    count = obj.count();
    count1 = obj1.count();
    ratio = (count / count1) * 100
    if ratio != 0:
        detection_ratio.objects.create(names=kword, ratio=ratio)

    ratio1 = ""
    kword1 = 'gender'
    print(kword1)
    obj1 = Cyberbullying_Detection_Type.objects.all().filter(Q(Prediction=kword1))
    obj11 = Cyberbullying_Detection_Type.objects.all()
    count1 = obj1.count();
    count11 = obj11.count();
    ratio1 = (count1 / count11) * 100
    if ratio1 != 0:
        detection_ratio.objects.create(names=kword1, ratio=ratio1)

    ratio12 = ""
    kword12 = 'religion'
    print(kword12)
    obj12 = Cyberbullying_Detection_Type.objects.all().filter(Q(Prediction=kword12))
    obj112 = Cyberbullying_Detection_Type.objects.all()
    count12 = obj12.count();
    count112 = obj112.count();
    ratio12 = (count12 / count112) * 100
    if ratio12 != 0:
        detection_ratio.objects.create(names=kword12, ratio=ratio12)

    ratio123 = ""
    kword123 = 'other_cyberbullying'
    print(kword123)
    obj123 = Cyberbullying_Detection_Type.objects.all().filter(Q(Prediction=kword123))
    obj1123 = Cyberbullying_Detection_Type.objects.all()
    count123 = obj123.count();
    count1123 = obj1123.count();
    ratio123 = (count123 / count1123) * 100
    if ratio123 != 0:
        detection_ratio.objects.create(names=kword123, ratio=ratio123)

    ratio1234 = ""
    kword1234 = 'age'
    print(kword1234)
    obj1234 = Cyberbullying_Detection_Type.objects.all().filter(Q(Prediction=kword1234))
    obj11234 = Cyberbullying_Detection_Type.objects.all()
    count1234 = obj1234.count();
    count11234 = obj11234.count();
    ratio1234 = (count1234 / count11234) * 100
    if ratio1234 != 0:
        detection_ratio.objects.create(names=kword1234, ratio=ratio1234)

    ratio123491 = ""
    kword123491 = 'ethnicity'
    print(kword123491)
    obj123491 = Cyberbullying_Detection_Type.objects.all().filter(Q(Prediction=kword123491))
    obj1123491 = Cyberbullying_Detection_Type.objects.all()
    count123491 = obj123491.count();
    count1123491 = obj1123491.count();
    ratio123491 = (count123491 / count1123491) * 100
    if ratio123491 != 0:
        detection_ratio.objects.create(names=kword123491, ratio=ratio123491)

    obj = detection_ratio.objects.all()
    return render(request, 'SProvider/Find_Predicted_Cyberbullying_Detection_Ratio.html', {'objs': obj})

def View_Remote_Users(request):
    obj=ClientRegister_Model.objects.all()
    return render(request,'SProvider/View_Remote_Users.html',{'objects':obj})

def ViewTrendings(request):
    topic = Cyberbullying_Detection_Type.objects.values('topics').annotate(dcount=Count('topics')).order_by('-dcount')
    return  render(request,'SProvider/ViewTrendings.html',{'objects':topic})

def charts(request,chart_type):
    chart1 = detection_ratio.objects.values('names').annotate(dcount=Avg('ratio'))
    return render(request,"SProvider/charts.html", {'form':chart1, 'chart_type':chart_type})

def charts1(request,chart_type):
    chart1 = detection_accuracy.objects.values('names').annotate(dcount=Avg('ratio'))
    return render(request,"SProvider/charts1.html", {'form':chart1, 'chart_type':chart_type})

def View_Predicted_Cyberbullying_Detection_Type(request):
    obj =Cyberbullying_Detection_Type.objects.all()
    return render(request, 'SProvider/View_Predicted_Cyberbullying_Detection_Type.html', {'list_objects': obj})

def likeschart(request,like_chart):
    charts =detection_accuracy.objects.values('names').annotate(dcount=Avg('ratio'))
    return render(request,"SProvider/likeschart.html", {'form':charts, 'like_chart':like_chart})


def Download_Predicted_DataSets(request):

    response = HttpResponse(content_type='application/ms-excel')
    # decide file name
    response['Content-Disposition'] = 'attachment; filename="Predicted_Data.xls"'
    # creating workbook
    wb = xlwt.Workbook(encoding='utf-8')
    # adding sheet
    ws = wb.add_sheet("sheet1")
    # Sheet header, first row
    row_num = 0
    font_style = xlwt.XFStyle()
    # headers are bold
    font_style.font.bold = True
    # writer = csv.writer(response)
    obj = Cyberbullying_Detection_Type.objects.all()
    data = obj  # dummy method to fetch data.
    for my_row in data:
        row_num = row_num + 1

        ws.write(row_num, 0, my_row.Tweet_Message, font_style)
        ws.write(row_num, 1, my_row.Prediction, font_style)

    wb.save(response)
    return response

def Train_Test_DataSets(request):
    detection_accuracy.objects.all().delete()

    data = pd.read_csv("Datasets.csv",encoding='latin-1')

    def clean_text(text):

        text = text.lower()
        text = re.sub('\[.*?\]', '', text)
        text = re.sub('https?://\S+|www\.\S+', '', text)
        text = re.sub('<.*?>+', '', text)
        text = re.sub('[%s]' % re.escape(string.punctuation), '', text)
        text = re.sub('\n', '', text)
        text = re.sub('\w*\d\w*', '', text)
        return text

        data['text'] = data['tweet_text'].apply(lambda x: clean_text(x))

    def apply_results(results):
        if (results == "not_cyberbullying"):
            return 0
        elif (results == "gender"):
            return 1
        elif (results == "religion"):
            return 2
        elif (results == "other_cyberbullying"):
            return 3
        elif (results == "age"):
            return 4
        elif (results == "ethnicity"):
            return 5

    data['Results'] = data['cyberbullying_type'].apply(apply_results)

    x = data['tweet_text']
    y = data['Results']

    cv = CountVectorizer(lowercase=False, strip_accents='unicode', ngram_range=(1, 1))

    x = cv.fit_transform(x)

    models = []
    from sklearn.model_selection import train_test_split
    X_train, X_test, y_train, y_test = train_test_split(x, y, test_size=0.20)
    X_train.shape, X_test.shape, y_train.shape

    print("Multinomial Naive Bayes")
    from sklearn.naive_bayes import MultinomialNB

    nb_clf = MultinomialNB()

    nb_clf.fit(X_train, y_train)
    MultinomialNB()
    nb_pred = nb_clf.predict(X_test)
    mnb = accuracy_score(y_test, nb_pred) * 100
    print(mnb)
    print(confusion_matrix(y_test, nb_pred))
    print(classification_report(y_test, nb_pred))
    models.append(('nb_pred', nb_clf))
    detection_accuracy.objects.create(names="MultinomialNB", ratio=mnb)


    # SVM Model
    print("SVM")
    from sklearn import svm
    lin_clf = svm.LinearSVC()
    lin_clf.fit(X_train, y_train)
    predict_svm = lin_clf.predict(X_test)
    svm_acc = accuracy_score(y_test, predict_svm) * 100
    print(svm_acc)
    print("CLASSIFICATION REPORT")
    print(classification_report(y_test, predict_svm))
    print("CONFUSION MATRIX")
    print(confusion_matrix(y_test, predict_svm))
    models.append(('svm', lin_clf))
    detection_accuracy.objects.create(names="SVM", ratio=svm_acc)


    print("Logistic Regression")
    from sklearn.linear_model import LogisticRegression
    reg = LogisticRegression(random_state=0, solver='lbfgs').fit(X_train, y_train)
    y_pred = reg.predict(X_test)
    print("ACCURACY")
    print(accuracy_score(y_test, y_pred) * 100)
    print("CLASSIFICATION REPORT")
    print(classification_report(y_test, y_pred))
    print("CONFUSION MATRIX")
    print(confusion_matrix(y_test, y_pred))
    models.append(('logistic', reg))
    detection_accuracy.objects.create(names="Logistic Regression", ratio=accuracy_score(y_test, y_pred) * 100)


    print("Decision Tree Classifier")
    dtc = DecisionTreeClassifier()
    dtc.fit(X_train, y_train)
    dtcpredict = dtc.predict(X_test)
    print("ACCURACY")
    print(accuracy_score(y_test, dtcpredict) * 100)
    print("CLASSIFICATION REPORT")
    print(classification_report(y_test, dtcpredict))
    print("CONFUSION MATRIX")
    print(confusion_matrix(y_test, dtcpredict))
    models.append(('DecisionTreeClassifier', dtc))
    detection_accuracy.objects.create(names="Decision Tree Classifier", ratio=accuracy_score(y_test, dtcpredict) * 100)


    print("SGD Classifier")
    from sklearn.linear_model import SGDClassifier
    sgd_clf = SGDClassifier(loss='hinge', penalty='l2', random_state=0)
    sgd_clf.fit(X_train, y_train)
    sgdpredict = sgd_clf.predict(X_test)
    print("ACCURACY")
    print(accuracy_score(y_test, sgdpredict) * 100)
    print("CLASSIFICATION REPORT")
    print(classification_report(y_test, sgdpredict))
    print("CONFUSION MATRIX")
    print(confusion_matrix(y_test, sgdpredict))
    models.append(('SGDClassifier', sgd_clf))
    detection_accuracy.objects.create(names="SGD Classifier", ratio=accuracy_score(y_test, sgdpredict) * 100)

    csv_format = 'Results.csv'
    data.to_csv(csv_format, index=False)
    data.to_markdown

    obj = detection_accuracy.objects.all()
    return render(request,'SProvider/Train_Test_DataSets.html', {'objs': obj})